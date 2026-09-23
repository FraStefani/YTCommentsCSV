"""Estrae tutti i commenti di un video YouTube e li salva in un report .xlsx colorato.

Uso tipico:
    python ytcomments.py "https://youtu.be/ByZrTRph2-4"
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

from xlsx_report import build_output_path, format_int, write_report
from yt_api import (
    ApiKeyError,
    Comment,
    CommentsDisabledError,
    InvalidUrlError,
    QuotaExceededError,
    VideoNotFoundError,
    YouTubeClient,
    YouTubeError,
    parse_video_id,
)

EXIT_OK = 0
EXIT_ERROR = 1
EXIT_VIDEO = 2
EXIT_KEY = 3
EXIT_QUOTA = 4
EXIT_PARTIAL = 5

PROGRESS_STEP = 200
_EPOCH = datetime.min.replace(tzinfo=timezone.utc)

MISSING_KEY_HELP = """ERRORE: API key di YouTube non trovata.

Come ottenerla (gratis, 2 minuti):
  1. apri https://console.cloud.google.com/ e crea un progetto
  2. "API e servizi" > "Libreria" > cerca "YouTube Data API v3" > Abilita
  3. "API e servizi" > "Credenziali" > "Crea credenziali" > "Chiave API"
  4. copia .env.example in .env e incolla la chiave dopo YOUTUBE_API_KEY=

In alternativa: variabile d'ambiente YOUTUBE_API_KEY, oppure opzione --api-key."""

EPILOG = """esempi:
  python ytcomments.py "https://youtu.be/ByZrTRph2-4"
  python ytcomments.py "https://youtu.be/ByZrTRph2-4" --max 50
  python ytcomments.py ByZrTRph2-4 --no-replies --out C:\\report

codici di uscita: 0 ok, 1 errore, 2 video/URL non valido, 3 API key,
4 quota esaurita, 5 report parziale."""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ytcomments.py",
        description="Scarica i commenti di un video YouTube in un file Excel colorato.",
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("url", help="link del video (o ID di 11 caratteri)")
    parser.add_argument("--max", type=int, metavar="N", default=None,
                        help="ferma lo scaricamento dopo N commenti (default: tutti)")
    parser.add_argument("--no-replies", action="store_true",
                        help="estrae solo i commenti di primo livello")
    parser.add_argument("--order", choices=("time", "relevance"), default="time",
                        help="ordine di scaricamento dall'API (default: time)")
    parser.add_argument("--out", metavar="DIR", default=None,
                        help="cartella di destinazione (default: quella dello script)")
    parser.add_argument("--name", default=None,
                        help="nome file senza estensione (default: canale_data)")
    parser.add_argument("--utc", action="store_true",
                        help="scrive le date in UTC invece che nell'ora locale")
    parser.add_argument("--api-key", default=None, help="API key di YouTube Data API v3")
    return parser


def read_env_file(path: Path) -> dict[str, str]:
    """Mini-parser di file .env: righe CHIAVE=valore, commenti con #."""
    values: dict[str, str] = {}
    try:
        raw = path.read_text(encoding="utf-8")
    except OSError:
        return values
    for line in raw.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        values[key.strip()] = value.strip().strip('"').strip("'")
    return values


def resolve_api_key(explicit: str | None = None) -> str:
    """Ordine di ricerca: opzione CLI, variabile d'ambiente, file .env accanto allo script."""
    if explicit and explicit.strip():
        return explicit.strip()
    from_env = os.environ.get("YOUTUBE_API_KEY", "").strip()
    if from_env:
        return from_env
    env_file = Path(__file__).resolve().parent / ".env"
    return read_env_file(env_file).get("YOUTUBE_API_KEY", "").strip()


def setup_console() -> None:
    """Evita UnicodeEncodeError con emoji e accenti nei titoli su console Windows (cp1252)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):  # pragma: no cover - stream non riconfigurabile
            pass


def _progress(count: int) -> None:
    print(f"\r  scaricati {format_int(count)} commenti...", end="", file=sys.stderr, flush=True)


def collect_comments(client: YouTubeClient, video_id: str, *, include_replies: bool,
                     order: str, limit: int | None) -> tuple[list[Comment], str, bool]:
    """Scarica i commenti restituendo (elenco ordinato, nota per il report, parziale?).

    Gli errori recuperabili (quota, rete, interruzione) non fanno perdere il lavoro
    già svolto: vengono annotati e i commenti raccolti fino a quel punto sono restituiti.
    """
    threads: list[tuple[Comment, list[Comment]]] = []
    note = ""
    partial = False
    total = 0
    next_milestone = PROGRESS_STEP
    try:
        for top, replies in client.iter_threads(
            video_id, include_replies=include_replies, order=order
        ):
            threads.append((top, replies))
            total += 1 + len(replies)
            if total >= next_milestone:
                _progress(total)
                next_milestone = total + PROGRESS_STEP
            if limit is not None and total >= limit:
                break
    except CommentsDisabledError as exc:
        note = f"{exc} Nessun commento da estrarre."
    except QuotaExceededError as exc:
        note = f"ATTENZIONE, report PARZIALE: {exc}"
        partial = True
    except KeyboardInterrupt:
        note = "ATTENZIONE, report PARZIALE: scaricamento interrotto manualmente."
        partial = True
    except YouTubeError as exc:
        note = f"ATTENZIONE, report PARZIALE: errore durante lo scaricamento ({exc})."
        partial = True
    finally:
        if total:
            _progress(total)
            print(file=sys.stderr)

    # Ordine cronologico crescente, con le risposte subito sotto il proprio commento.
    threads.sort(key=lambda pair: pair[0].published_at or _EPOCH)
    flat: list[Comment] = []
    for top, replies in threads:
        flat.append(top)
        flat.extend(sorted(replies, key=lambda item: item.published_at or _EPOCH))

    if limit is not None and len(flat) > limit:
        flat = flat[:limit]
        extra = f"Elenco limitato ai primi {format_int(limit)} commenti (opzione --max)."
        note = f"{note} {extra}".strip()
    return flat, note, partial


def main(argv: list[str] | None = None) -> int:
    setup_console()
    args = build_parser().parse_args(argv)

    if args.max is not None and args.max <= 0:
        print("ERRORE: --max deve essere maggiore di zero.", file=sys.stderr)
        return EXIT_ERROR

    try:
        video_id = parse_video_id(args.url)
    except InvalidUrlError as exc:
        print(f"ERRORE: {exc}", file=sys.stderr)
        return EXIT_VIDEO

    api_key = resolve_api_key(args.api_key)
    if not api_key:
        print(MISSING_KEY_HELP, file=sys.stderr)
        return EXIT_KEY

    client = YouTubeClient(api_key)
    try:
        video = client.get_video_info(video_id)
    except VideoNotFoundError as exc:
        print(f"ERRORE: {exc}", file=sys.stderr)
        return EXIT_VIDEO
    except ApiKeyError as exc:
        print(f"ERRORE: {exc}", file=sys.stderr)
        return EXIT_KEY
    except QuotaExceededError as exc:
        print(f"ERRORE: {exc}", file=sys.stderr)
        return EXIT_QUOTA
    except YouTubeError as exc:
        print(f"ERRORE: {exc}", file=sys.stderr)
        return EXIT_ERROR

    print(f"Canale ......... {video.channel}")
    print(f"Video .......... {video.title}")
    print(f"Visualizzazioni  {format_int(video.view_count)}")
    print(f"Commenti ....... {format_int(video.comment_count)} (dichiarati da YouTube)")

    comments, note, partial = collect_comments(
        client,
        video_id,
        include_replies=not args.no_replies,
        order=args.order,
        limit=args.max,
    )

    out_dir = Path(args.out).expanduser() if args.out else Path(__file__).resolve().parent
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        print(f"ERRORE: cartella di destinazione non utilizzabile ({exc}).", file=sys.stderr)
        return EXIT_ERROR

    extracted_at = datetime.now(timezone.utc)
    path = build_output_path(video.channel, datetime.now(), out_dir, args.name)
    try:
        saved = write_report(
            path, video, comments, extracted_at=extracted_at, note=note, utc=args.utc
        )
    except PermissionError:
        print(
            f"ERRORE: impossibile scrivere in {path.parent}. "
            "Chiudi il file in Excel oppure usa --out per un'altra cartella.",
            file=sys.stderr,
        )
        return EXIT_ERROR

    replies = sum(1 for item in comments if item.is_reply)
    print(f"\nEstratti ....... {format_int(len(comments))} "
          f"({format_int(len(comments) - replies)} commenti + {format_int(replies)} risposte)")
    print(f"Quota usata .... {client.calls} unità (chiamate API)")
    print(f"File ........... {saved}")
    if note:
        print(f"\n{note}")
    return EXIT_PARTIAL if partial else EXIT_OK


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\nInterrotto prima di scaricare i commenti.", file=sys.stderr)
        sys.exit(EXIT_PARTIAL)

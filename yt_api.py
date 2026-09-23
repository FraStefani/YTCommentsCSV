"""Accesso a YouTube Data API v3 usando solo la libreria standard di Python.

Il modulo espone il parsing dell'URL del video, un client HTTP con retry e la
paginazione di commenti e risposte. Nessuna dipendenza esterna: urllib + json.
"""

from __future__ import annotations

import html
import json
import random
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import datetime, timezone

API_BASE = "https://www.googleapis.com/youtube/v3/"
PAGE_SIZE = 100
USER_AGENT = "YTCommentsCSV/1.0 (+stdlib urllib)"

_VIDEO_ID_RE = re.compile(r"^[A-Za-z0-9_-]{11}$")
_PATH_PREFIXES = {"shorts", "live", "embed", "v"}
_ENTITY_RE = re.compile(r"&(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);")
_RETRY_STATUS = {429, 500, 502, 503, 504}
_RETRY_REASONS = {"rateLimitExceeded", "userRateLimitExceeded", "backendError", "internalError"}


class YouTubeError(Exception):
    """Errore generico nell'accesso a YouTube."""


class InvalidUrlError(YouTubeError):
    """L'URL fornito non contiene un ID video valido."""


class VideoNotFoundError(YouTubeError):
    """Il video non esiste, è privato oppure è stato rimosso."""


class CommentsDisabledError(YouTubeError):
    """I commenti sono disabilitati sul video."""


class QuotaExceededError(YouTubeError):
    """Quota giornaliera dell'API esaurita."""


class ApiKeyError(YouTubeError):
    """API key assente, non valida o non autorizzata."""


@dataclass
class VideoInfo:
    """Metadati del video mostrati nella tabella iniziale del report."""

    video_id: str
    url: str
    title: str
    channel: str
    published_at: datetime | None = None
    view_count: int | None = None
    comment_count: int | None = None


@dataclass
class Comment:
    """Un commento di primo livello o una risposta."""

    author: str
    text: str
    published_at: datetime | None = None
    like_count: int = 0
    comment_id: str = ""
    is_reply: bool = False
    parent_author: str = ""


def canonical_url(video_id: str) -> str:
    """URL pulito del video, senza parametri di tracciamento."""
    return f"https://www.youtube.com/watch?v={video_id}"


def parse_video_id(url: str) -> str:
    """Ricava l'ID video dai formati più comuni di link YouTube.

    Gestisce youtu.be/ID, watch?v=ID, /shorts/ID, /live/ID, /embed/ID, /v/ID,
    parametri extra (si=, t=, list=) e l'ID nudo di 11 caratteri.
    """
    raw = (url or "").strip().strip('"').strip("'")
    if not raw:
        raise InvalidUrlError("Nessun URL fornito.")
    if _VIDEO_ID_RE.match(raw):
        return raw
    if "://" not in raw:
        raw = "https://" + raw

    parts = urllib.parse.urlsplit(raw)
    host = parts.netloc.lower().split(":")[0]
    for prefix in ("www.", "m.", "music."):
        if host.startswith(prefix):
            host = host[len(prefix):]
    segments = [seg for seg in parts.path.split("/") if seg]

    candidate: str | None = None
    if host == "youtu.be":
        candidate = segments[0] if segments else None
    elif host in {"youtube.com", "youtube-nocookie.com"}:
        query = urllib.parse.parse_qs(parts.query)
        if query.get("v"):
            candidate = query["v"][0]
        elif len(segments) >= 2 and segments[0] in _PATH_PREFIXES:
            candidate = segments[1]
    else:
        raise InvalidUrlError(f"Non sembra un link YouTube: {url!r}")

    if candidate and _VIDEO_ID_RE.match(candidate):
        return candidate
    raise InvalidUrlError(f"Impossibile ricavare l'ID video da: {url!r}")


def _as_int(value: object) -> int | None:
    try:
        return int(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return None


def _parse_ts(value: object) -> datetime | None:
    """Converte un timestamp RFC 3339 dell'API in datetime con timezone UTC."""
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)


def _comment_from(payload: dict, *, is_reply: bool, parent_author: str = "") -> Comment:
    """Costruisce un Comment da una risorsa `comment`, senza mai alzare KeyError."""
    snippet = (payload or {}).get("snippet") or {}
    author = (snippet.get("authorDisplayName") or "").strip() or "[utente rimosso]"
    # textOriginal è restituito solo all'autore autenticato: con la sola API key si
    # usa textDisplay, che con textFormat=plainText arriva senza markup HTML.
    text = snippet.get("textOriginal")
    if text is None:
        text = snippet.get("textDisplay") or ""
    text = str(text)
    if _ENTITY_RE.search(text):
        text = html.unescape(text)
    return Comment(
        author=author,
        text=text,
        published_at=_parse_ts(snippet.get("publishedAt")),
        like_count=_as_int(snippet.get("likeCount")) or 0,
        comment_id=(payload or {}).get("id") or "",
        is_reply=is_reply,
        parent_author=parent_author,
    )


def _read_api_error(exc: urllib.error.HTTPError) -> tuple[str, str]:
    """Estrae (reason, message) dal corpo JSON di un errore dell'API."""
    try:
        body = json.loads(exc.read().decode("utf-8", "replace"))
    except Exception:
        return "", ""
    error = body.get("error") or {}
    errors = error.get("errors") or [{}]
    return errors[0].get("reason", "") or "", error.get("message", "") or ""


def _map_error(status: int, reason: str, message: str) -> YouTubeError | None:
    """Traduce un errore HTTP in eccezione specifica, o None se conviene riprovare."""
    if reason == "commentsDisabled":
        return CommentsDisabledError("I commenti sono disabilitati su questo video.")
    if reason in {"quotaExceeded", "dailyLimitExceeded"}:
        return QuotaExceededError(
            "Quota giornaliera dell'API esaurita (si azzera a mezzanotte, ora del Pacifico)."
        )
    if reason in {"keyInvalid", "keyExpired", "ipRefererBlocked", "forbidden"} or "API key not valid" in message:
        return ApiKeyError(f"API key rifiutata da Google: {message or reason}")
    if reason in {"videoNotFound", "notFound"}:
        return VideoNotFoundError(
            "Video non trovato: potrebbe essere privato, rimosso, oppure l'ID è errato."
        )
    if status == 400:
        return YouTubeError(f"Richiesta rifiutata dall'API ({reason or status}): {message}")
    return None


class YouTubeClient:
    """Client con retry esponenziale e conteggio chiamate (1 chiamata = 1 unità di quota)."""

    def __init__(self, api_key: str, *, timeout: float = 30.0, max_retries: int = 5) -> None:
        self.api_key = api_key
        self.timeout = timeout
        self.max_retries = max_retries
        self.calls = 0

    def _get(self, endpoint: str, **params: object) -> dict:
        query = {key: value for key, value in params.items() if value is not None}
        query["key"] = self.api_key
        url = API_BASE + endpoint + "?" + urllib.parse.urlencode(query)
        # Nota: l'URL (che contiene la key) non finisce mai nei messaggi di errore.
        delay = 1.0
        for attempt in range(1, self.max_retries + 1):
            try:
                self.calls += 1
                request = urllib.request.Request(
                    url, headers={"Accept": "application/json", "User-Agent": USER_AGENT}
                )
                with urllib.request.urlopen(request, timeout=self.timeout) as response:
                    return json.loads(response.read().decode("utf-8"))
            except urllib.error.HTTPError as exc:
                reason, message = _read_api_error(exc)
                mapped = _map_error(exc.code, reason, message)
                if mapped is not None:
                    raise mapped from None
                if exc.code in _RETRY_STATUS or reason in _RETRY_REASONS:
                    if attempt == self.max_retries:
                        raise YouTubeError(
                            f"L'API continua a rispondere con errore {exc.code} "
                            f"({reason or 'nessun dettaglio'}) dopo {attempt} tentativi."
                        ) from None
                    time.sleep(delay + random.uniform(0, 0.5))
                    delay *= 2
                    continue
                raise YouTubeError(
                    f"Errore HTTP {exc.code} dall'API ({reason or 'nessun dettaglio'}): {message}"
                ) from None
            except (urllib.error.URLError, TimeoutError) as exc:
                if attempt == self.max_retries:
                    raise YouTubeError(
                        "Impossibile raggiungere googleapis.com: verifica la connessione o il "
                        f"proxy aziendale (HTTPS_PROXY). Dettaglio: {exc}"
                    ) from None
                time.sleep(delay + random.uniform(0, 0.5))
                delay *= 2
        raise YouTubeError("Chiamata API non riuscita.")  # pragma: no cover

    def get_video_info(self, video_id: str) -> VideoInfo:
        """Metadati del video: canale, titolo, data di uscita, visualizzazioni, n. commenti."""
        data = self._get("videos", part="snippet,statistics", id=video_id)
        items = data.get("items") or []
        if not items:
            raise VideoNotFoundError(
                f"Nessun video con ID {video_id}: potrebbe essere privato, rimosso "
                "oppure l'URL è errato."
            )
        snippet = items[0].get("snippet") or {}
        stats = items[0].get("statistics") or {}
        return VideoInfo(
            video_id=video_id,
            url=canonical_url(video_id),
            title=snippet.get("title") or "[titolo non disponibile]",
            channel=snippet.get("channelTitle") or "[canale non disponibile]",
            published_at=_parse_ts(snippet.get("publishedAt")),
            view_count=_as_int(stats.get("viewCount")),
            comment_count=_as_int(stats.get("commentCount")),
        )

    def _iter_replies(self, parent_id: str, parent_author: str):
        """Tutte le risposte a un commento (comments.list paginato)."""
        page_token = None
        while True:
            data = self._get(
                "comments",
                part="snippet",
                parentId=parent_id,
                maxResults=PAGE_SIZE,
                textFormat="plainText",
                pageToken=page_token,
            )
            for item in data.get("items") or []:
                yield _comment_from(item, is_reply=True, parent_author=parent_author)
            page_token = data.get("nextPageToken")
            if not page_token:
                return

    def iter_threads(self, video_id: str, *, include_replies: bool = True, order: str = "time"):
        """Genera coppie (commento di primo livello, risposte) paginando commentThreads."""
        page_token = None
        while True:
            data = self._get(
                "commentThreads",
                part="snippet,replies",
                videoId=video_id,
                maxResults=PAGE_SIZE,
                order=order,
                textFormat="plainText",
                pageToken=page_token,
            )
            for item in data.get("items") or []:
                thread = item.get("snippet") or {}
                top = _comment_from(thread.get("topLevelComment") or {}, is_reply=False)
                replies: list[Comment] = []
                expected = _as_int(thread.get("totalReplyCount")) or 0
                if include_replies and expected:
                    inline = ((item.get("replies") or {}).get("comments")) or []
                    if len(inline) >= expected:
                        # Le risposte inline bastano: nessuna chiamata extra, quota risparmiata.
                        replies = [
                            _comment_from(reply, is_reply=True, parent_author=top.author)
                            for reply in inline
                        ]
                    else:
                        # La parte `replies` è parziale per design: serve comments.list.
                        replies = list(self._iter_replies(top.comment_id, top.author))
                yield top, replies
            page_token = data.get("nextPageToken")
            if not page_token:
                return

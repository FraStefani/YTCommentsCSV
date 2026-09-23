"""Costruzione del report .xlsx colorato: due tabelle, metadati in alto e commenti sotto.

Il layout non usa celle unite, così lo stesso codice funziona sia in modalità normale
sia in modalità write_only (necessaria sopra le decine di migliaia di commenti, dove
merge_cells non è disponibile).
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from yt_api import Comment, VideoInfo

# --- Palette -----------------------------------------------------------------
NAVY = "1F3864"          # bande titolo e intestazioni di tabella
BLUE_LABEL = "2E5C8A"    # etichette dei metadati
BLUE_VALUE = "EAF2FB"    # valori dei metadati
ROW_A = "FFFFFF"         # righe commento, alternate
ROW_B = "F4F8FD"
REPLY_A = "FFF6E5"       # righe risposta, alternate
REPLY_B = "FFEFD4"
BORDER_COLOR = "C9D6E8"
WHITE = "FFFFFF"
INK = "1A1A1A"
INK_DARK = "14274E"
INK_REPLY = "5C4400"
LINK_COLOR = "0563C1"

# --- Struttura della tabella commenti ----------------------------------------
COLUMNS: tuple[tuple[str, int], ...] = (
    ("Utente", 26),
    ("Commento", 95),
    ("Data commento", 18),
    ("Like", 8),
    ("Tipo", 12),
    ("Risposta a", 24),
)
N_COLS = len(COLUMNS)

DATE_FORMAT = "DD/MM/YYYY HH:MM"
NUMBER_FORMAT = "#,##0"
MAX_CELL_CHARS = 32767          # limite fisico di Excel
TEXT_LIMIT = 32000              # margine di sicurezza per il testo dei commenti
TRUNCATE_MARK = "…[troncato]"
WRITE_ONLY_THRESHOLD = 50_000   # oltre questa soglia si scrive in streaming

_SIDE = Side(style="thin", color=BORDER_COLOR)
BORDER = Border(left=_SIDE, right=_SIDE, top=_SIDE, bottom=_SIDE)

_FONT_TITLE = Font(bold=True, size=14, color=WHITE)
_FONT_SECTION = Font(bold=True, size=11, color=WHITE)
_FONT_LABEL = Font(bold=True, color=WHITE)
_FONT_VALUE = Font(color=INK_DARK)
_FONT_LINK = Font(color=LINK_COLOR, underline="single")
_FONT_HEADER = Font(bold=True, color=WHITE)
_FONT_COMMENT = Font(color=INK)
_FONT_REPLY = Font(color=INK_REPLY)

_ALIGN_BAND = Alignment(horizontal="left", vertical="center")
_ALIGN_LABEL = Alignment(horizontal="left", vertical="top", wrap_text=True)
_ALIGN_TEXT = Alignment(horizontal="left", vertical="top", wrap_text=True)
_ALIGN_PLAIN = Alignment(horizontal="left", vertical="top")
_ALIGN_NUM = Alignment(horizontal="right", vertical="top")
_ALIGN_TEXT_REPLY = Alignment(horizontal="left", vertical="top", wrap_text=True, indent=2)
_ALIGN_PLAIN_REPLY = Alignment(horizontal="left", vertical="top", indent=2)

_INVALID_FILENAME_RE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_RESERVED_NAMES = {"CON", "PRN", "AUX", "NUL"} | {
    f"{prefix}{i}" for prefix in ("COM", "LPT") for i in range(1, 10)
}


@dataclass
class _Spec:
    """Una cella da scrivere: valore più stile, indipendente dalla modalità del foglio."""

    value: object = None
    fill: PatternFill | None = None
    font: Font | None = None
    alignment: Alignment | None = None
    number_format: str | None = None
    hyperlink: str | None = None
    force_text: bool = False
    border: bool = False


# --- Utilità di sanitizzazione ------------------------------------------------
def clean_text(value: object, limit: int = TEXT_LIMIT) -> str:
    """Rende un testo scrivibile in xlsx: niente caratteri illegali, niente CR, lunghezza limitata."""
    text = "" if value is None else str(value)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = ILLEGAL_CHARACTERS_RE.sub("", text)
    limit = min(limit, MAX_CELL_CHARS)
    if len(text) > limit:
        text = text[: limit - len(TRUNCATE_MARK)] + TRUNCATE_MARK
    return text


def to_excel_datetime(value: datetime | None, utc: bool = False) -> datetime | None:
    """Excel non supporta i fusi orari: converte e rimuove tzinfo."""
    if value is None:
        return None
    if value.tzinfo is None:
        return value.replace(microsecond=0)
    converted = value.astimezone(timezone.utc) if utc else value.astimezone()
    return converted.replace(tzinfo=None, microsecond=0)


def format_int(value: int | None) -> str:
    """Numero con separatore delle migliaia in stile italiano."""
    if value is None:
        return "N/D"
    return f"{value:,}".replace(",", ".")


def sanitize_filename(name: str, max_len: int = 150) -> str:
    """Nome file valido su Windows: niente caratteri vietati, nomi riservati o punti finali."""
    cleaned = _INVALID_FILENAME_RE.sub("_", (name or "").strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].strip(" .")
    if not cleaned:
        cleaned = "canale_sconosciuto"
    if cleaned.split(".")[0].upper() in _RESERVED_NAMES:
        cleaned = f"_{cleaned}"
    return cleaned


def build_output_path(channel: str, when: datetime, out_dir: Path, override: str | None = None) -> Path:
    """<nome canale>_<data attuale>.xlsx, con suffisso progressivo se il file esiste già."""
    if override and override.lower().endswith(".xlsx"):
        override = override[: -len(".xlsx")]
    base = sanitize_filename(override or f"{channel}_{when:%Y-%m-%d}")
    out_dir = Path(out_dir)
    candidate = out_dir / f"{base}.xlsx"
    counter = 2
    while candidate.exists():
        candidate = out_dir / f"{base}_{counter}.xlsx"
        counter += 1
    return candidate


# --- Costruzione delle righe ---------------------------------------------------
def _band(text: str, fill_color: str, font: Font) -> list[_Spec]:
    """Banda colorata su tutta la larghezza della tabella, testo nella prima cella."""
    fill = PatternFill("solid", fgColor=fill_color)
    return [
        _Spec(
            value=clean_text(text) if idx == 0 else None,
            fill=fill,
            font=font,
            alignment=_ALIGN_BAND,
            force_text=idx == 0,
        )
        for idx in range(N_COLS)
    ]


def _meta_row(label: str, value: object, number_format: str | None = None,
              hyperlink: str | None = None) -> list[_Spec]:
    """Etichetta colorata in colonna A, valore su sfondo chiaro in colonna B."""
    is_text = not isinstance(value, (int, float, datetime)) or isinstance(value, bool)
    specs = [
        _Spec(
            value=clean_text(label),
            fill=PatternFill("solid", fgColor=BLUE_LABEL),
            font=_FONT_LABEL,
            alignment=_ALIGN_LABEL,
            force_text=True,
            border=True,
        ),
        _Spec(
            value=clean_text(value) if is_text else value,
            fill=PatternFill("solid", fgColor=BLUE_VALUE),
            font=_FONT_LINK if hyperlink else _FONT_VALUE,
            alignment=_ALIGN_LABEL,
            number_format=number_format,
            hyperlink=hyperlink,
            force_text=is_text,
            border=True,
        ),
    ]
    # Le colonne restanti restano vuote e senza sfondo, così i valori lunghi
    # (link, titolo) possono sbordare visivamente senza essere tagliati.
    specs.extend(_Spec() for _ in range(N_COLS - 2))
    return specs


def _header_row() -> list[_Spec]:
    fill = PatternFill("solid", fgColor=NAVY)
    return [
        _Spec(value=title, fill=fill, font=_FONT_HEADER, alignment=_ALIGN_BAND,
              force_text=True, border=True)
        for title, _ in COLUMNS
    ]


def _comment_row(comment: Comment, index: int, utc: bool) -> list[_Spec]:
    """Una riga di commento o di risposta, con banding alternato e rientro per le risposte."""
    alternate = index % 2 == 1
    if comment.is_reply:
        fill = PatternFill("solid", fgColor=REPLY_B if alternate else REPLY_A)
        font = _FONT_REPLY
        align_text, align_plain = _ALIGN_TEXT_REPLY, _ALIGN_PLAIN_REPLY
    else:
        fill = PatternFill("solid", fgColor=ROW_B if alternate else ROW_A)
        font = _FONT_COMMENT
        align_text, align_plain = _ALIGN_TEXT, _ALIGN_PLAIN

    def cell(value: object, alignment: Alignment, number_format: str | None = None,
             force_text: bool = True) -> _Spec:
        return _Spec(
            value=clean_text(value) if force_text else value,
            fill=fill,
            font=font,
            alignment=alignment,
            number_format=number_format,
            force_text=force_text,
            border=True,
        )

    return [
        cell(comment.author, align_plain),
        cell(comment.text, align_text),
        cell(to_excel_datetime(comment.published_at, utc), _ALIGN_PLAIN,
             DATE_FORMAT, force_text=False),
        cell(comment.like_count, _ALIGN_NUM, NUMBER_FORMAT, force_text=False),
        cell("Risposta" if comment.is_reply else "Commento", align_plain),
        cell(comment.parent_author if comment.is_reply else "", align_plain),
    ]


def _metadata(video: VideoInfo, comments: list[Comment], extracted_at: datetime,
              note: str, utc: bool) -> list[list[_Spec]]:
    replies = sum(1 for item in comments if item.is_reply)
    tops = len(comments) - replies
    extracted = (
        f"{format_int(len(comments))}  ({format_int(tops)} commenti + {format_int(replies)} risposte)"
    )
    rows = [
        _meta_row("Link video", video.url, hyperlink=video.url),
        _meta_row("Canale", video.channel),
        _meta_row("Titolo video", video.title),
        _meta_row("Data uscita video", to_excel_datetime(video.published_at, utc), DATE_FORMAT),
        _meta_row(
            "Visualizzazioni",
            video.view_count if video.view_count is not None else "N/D",
            NUMBER_FORMAT if video.view_count is not None else None,
        ),
        _meta_row("Commenti dichiarati da YouTube", format_int(video.comment_count)),
        _meta_row("Commenti estratti", extracted),
        _meta_row("Estrazione del", to_excel_datetime(extracted_at, utc), DATE_FORMAT),
    ]
    if note:
        rows.append(_meta_row("Note", note))
    return rows


# --- Scrittura ----------------------------------------------------------------
def _apply(cell, spec: _Spec) -> None:
    if spec.fill is not None:
        cell.fill = spec.fill
    if spec.font is not None:
        cell.font = spec.font
    if spec.alignment is not None:
        cell.alignment = spec.alignment
    if spec.number_format:
        cell.number_format = spec.number_format
    if spec.border:
        cell.border = BORDER
    if spec.hyperlink:
        try:
            cell.hyperlink = spec.hyperlink
        except Exception:  # pragma: no cover - il link resta come testo visibile
            pass
    # Neutralizza le formule: un commento che inizia con "=" resta testo.
    if spec.force_text and isinstance(spec.value, str) and spec.value:
        cell.data_type = "s"


def _emit(ws, row_index: int, specs: Iterable[_Spec], write_only: bool) -> None:
    if write_only:
        cells = []
        for spec in specs:
            cell = WriteOnlyCell(ws, value=spec.value)
            _apply(cell, spec)
            cells.append(cell)
        ws.append(cells)
    else:
        for column, spec in enumerate(specs, start=1):
            cell = ws.cell(row=row_index, column=column, value=spec.value)
            _apply(cell, spec)


def _save(workbook: Workbook, path: Path, attempts: int = 5) -> Path:
    """Salva gestendo il caso in cui il file sia bloccato (aperto in Excel)."""
    base = path
    for attempt in range(attempts):
        try:
            workbook.save(path)
            return path
        except PermissionError:
            if attempt == attempts - 1:
                raise
            path = base.with_name(f"{base.stem}_{attempt + 2}{base.suffix}")
    raise PermissionError(str(base))  # pragma: no cover


def write_report(path: Path, video: VideoInfo, comments: list[Comment], *,
                 extracted_at: datetime, note: str = "", utc: bool = False,
                 write_only: bool | None = None) -> Path:
    """Scrive il report e restituisce il percorso effettivo del file salvato."""
    if write_only is None:
        write_only = len(comments) > WRITE_ONLY_THRESHOLD

    workbook = Workbook(write_only=write_only)
    if write_only:
        sheet = workbook.create_sheet("Report")
    else:
        sheet = workbook.active
        sheet.title = "Report"

    for index, (_, width) in enumerate(COLUMNS, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width

    meta_rows = _metadata(video, comments, extracted_at, note, utc)
    # Layout: 1 banda titolo, 1 riga vuota, N metadati, 1 riga vuota, 1 banda sezione, 1 header.
    header_row = len(meta_rows) + 5
    last_row = header_row + max(len(comments), 1)
    last_column = get_column_letter(N_COLS)

    # Nessun blocco dei riquadri: il foglio scorre tutto, metadati compresi.
    # In write_only il filtro va impostato prima di scrivere le righe.
    sheet.auto_filter.ref = f"A{header_row}:{last_column}{last_row}"

    row = 0

    def emit(specs: list[_Spec]) -> None:
        nonlocal row
        row += 1
        _emit(sheet, row, specs, write_only)

    emit(_band("REPORT COMMENTI YOUTUBE", NAVY, _FONT_TITLE))
    emit([_Spec() for _ in range(N_COLS)])
    for meta in meta_rows:
        emit(meta)
    emit([_Spec() for _ in range(N_COLS)])
    emit(_band("COMMENTI", NAVY, _FONT_SECTION))
    emit(_header_row())

    if comments:
        for index, comment in enumerate(comments):
            emit(_comment_row(comment, index, utc))
    else:
        empty = _band("Nessun commento estratto.", ROW_B, _FONT_COMMENT)
        emit(empty)

    return _save(workbook, Path(path))

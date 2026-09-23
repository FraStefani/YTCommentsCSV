"""Test offline dell'orchestrazione: ordinamento, limiti e salvataggio parziale."""

from __future__ import annotations

import sys
import tempfile
import unittest
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from yt_api import Comment, CommentsDisabledError, QuotaExceededError
from ytcomments import collect_comments, read_env_file

BASE = datetime(2024, 3, 14, 12, 0, tzinfo=timezone.utc)


def _comment(author: str, minutes: int, *, is_reply: bool = False, parent: str = "") -> Comment:
    return Comment(
        author=author,
        text=f"testo di {author}",
        published_at=BASE + timedelta(minutes=minutes),
        comment_id=f"id-{author}",
        is_reply=is_reply,
        parent_author=parent,
    )


class _StubClient:
    """Finto client: restituisce thread predefiniti e opzionalmente esplode a metà."""

    def __init__(self, threads, error: Exception | None = None) -> None:
        self.threads = threads
        self.error = error
        self.calls = 0

    def iter_threads(self, video_id, *, include_replies=True, order="time"):
        for thread in self.threads:
            self.calls += 1
            yield thread
        if self.error is not None:
            raise self.error


class TestCollectComments(unittest.TestCase):
    def _threads(self):
        # Volutamente in ordine non cronologico, come li restituisce order=time.
        return [
            (_comment("carla", 30), []),
            (
                _comment("anna", 10),
                [
                    _comment("dario", 25, is_reply=True, parent="anna"),
                    _comment("bruno", 15, is_reply=True, parent="anna"),
                ],
            ),
        ]

    def test_ordine_cronologico_con_risposte_sotto_il_padre(self) -> None:
        client = _StubClient(self._threads())
        comments, note, partial = collect_comments(
            client, "vid", include_replies=True, order="time", limit=None
        )
        self.assertEqual([c.author for c in comments], ["anna", "bruno", "dario", "carla"])
        self.assertEqual(note, "")
        self.assertFalse(partial)

    def test_limite_max_tronca_ed_annota(self) -> None:
        client = _StubClient(self._threads())
        comments, note, partial = collect_comments(
            client, "vid", include_replies=True, order="time", limit=2
        )
        self.assertEqual(len(comments), 2)
        self.assertIn("--max", note)
        self.assertFalse(partial)

    def test_quota_esaurita_conserva_i_commenti_scaricati(self) -> None:
        client = _StubClient(self._threads(), error=QuotaExceededError("quota finita"))
        comments, note, partial = collect_comments(
            client, "vid", include_replies=True, order="time", limit=None
        )
        self.assertEqual(len(comments), 4)
        self.assertTrue(partial)
        self.assertIn("PARZIALE", note)

    def test_interruzione_manuale_conserva_i_commenti(self) -> None:
        client = _StubClient(self._threads(), error=KeyboardInterrupt())
        comments, note, partial = collect_comments(
            client, "vid", include_replies=True, order="time", limit=None
        )
        self.assertEqual(len(comments), 4)
        self.assertTrue(partial)
        self.assertIn("interrotto", note)

    def test_commenti_disabilitati_non_e_un_errore_fatale(self) -> None:
        client = _StubClient([], error=CommentsDisabledError("I commenti sono disabilitati."))
        comments, note, partial = collect_comments(
            client, "vid", include_replies=True, order="time", limit=None
        )
        self.assertEqual(comments, [])
        self.assertFalse(partial)
        self.assertIn("disabilitati", note)


class TestEnvFile(unittest.TestCase):
    def test_lettura_env(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            percorso = Path(tmp) / ".env"
            percorso.write_text(
                "# commento\n\nYOUTUBE_API_KEY = \"AIza-test\"\nALTRO=1\nriga-senza-uguale\n",
                encoding="utf-8",
            )
            valori = read_env_file(percorso)
        self.assertEqual(valori["YOUTUBE_API_KEY"], "AIza-test")
        self.assertEqual(valori["ALTRO"], "1")
        self.assertNotIn("riga-senza-uguale", valori)

    def test_file_assente(self) -> None:
        self.assertEqual(read_env_file(Path("nessun-file-qui.env")), {})


if __name__ == "__main__":
    unittest.main(verbosity=2)

"""Worker process tests: protocol, isolation and (when Word exists) real splits.

The Word tests are skipped automatically on machines without Microsoft Word
desktop, or when the environment variable WS_SKIP_WORD is set.
"""

from __future__ import annotations

import os
import queue
import sys
import tempfile
import time
import unittest
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC))

import i18n  # noqa: E402
import jobs  # noqa: E402
import sample_doc  # noqa: E402
import winsys  # noqa: E402


def word_installed() -> bool:
    if os.environ.get("WS_SKIP_WORD") or os.name != "nt":
        return False
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"):
            return True
    except OSError:
        return False


def run(request: dict, timeout: float = 300.0, cancel_on: str = ""):
    job = jobs.WorkerJob(request)
    final = None
    progress = []
    deadline = time.monotonic() + timeout
    while True:
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            job.cancel()
            raise AssertionError("worker timed out")
        try:
            kind, data = job.events.get(timeout=remaining)
        except queue.Empty:
            continue
        if kind == "progress":
            progress.append(data.get("key"))
            if cancel_on and data.get("key") == cancel_on:
                job.cancel()
        if kind in ("done", "error", "cancelled") and final is None:
            final = (kind, data)
        if kind == "exit":
            return final, progress, job


class WorkerProtocolTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_unknown_job_kind_is_reported(self):
        (kind, data), _progress, job = run({"kind": "nonsense", "source": "x"})
        self.assertEqual(kind, "error")
        self.assertEqual(data["error"]["key"], "err.unexpected")
        self.assertFalse(job.job_dir.exists())

    def test_invalid_document_fails_before_word(self):
        bogus = self.dir / "bogus.docx"
        bogus.write_text("plain text", encoding="utf-8")
        before = winsys.pids_by_image("WINWORD.EXE")
        (kind, data), progress, _job = run({"kind": "analyze", "source": str(bogus)})
        self.assertEqual(kind, "error")
        self.assertEqual(data["error"]["key"], "err.not_zip")
        self.assertNotIn("progress.starting_word", progress)
        self.assertLessEqual(winsys.pids_by_image("WINWORD.EXE") - before, set())


@unittest.skipUnless(word_installed(), "Microsoft Word desktop is not available")
class WordIntegrationTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.words_before = winsys.pids_by_image("WINWORD.EXE")
        tr = i18n.Translator("en")
        self.sample = sample_doc.build(self.dir / "Sample.docx", tr.t)

    def tearDown(self) -> None:
        # No automation Word may outlive a test (allow a moment for a clean quit).
        for _ in range(100):
            if not (winsys.pids_by_image("WINWORD.EXE") - self.words_before):
                break
            time.sleep(0.1)
        leftover = winsys.pids_by_image("WINWORD.EXE") - self.words_before
        self._tmp.cleanup()
        self.assertEqual(leftover, set(), "a Word process outlived the worker")

    def analyze(self):
        (kind, data), _p, _j = run({"kind": "analyze", "source": str(self.sample)})
        self.assertEqual(kind, "done", data)
        return data["result"]

    def test_sample_has_twelve_pages_and_splits_exactly(self):
        info = self.analyze()
        self.assertEqual(info["pages"], sample_doc.PAGES)
        request = {
            "kind": "split", "source": str(self.sample), "fingerprint": info["fingerprint"],
            "total_pages": info["pages"], "page_a": sample_doc.SPLIT_A, "page_b": sample_doc.SPLIT_B,
            "output_one": str(self.dir / "One.docx"), "output_two": str(self.dir / "Two.docx"),
        }
        (kind, data), progress, _job = run(request)
        self.assertEqual(kind, "done", data)
        result = data["result"]
        self.assertEqual((result["pages_one"], result["pages_two"]), (6, 6))
        self.assertEqual(result["warnings"], [])
        self.assertIn("progress.verifying", progress)
        names = sorted(p.name for p in self.dir.iterdir())
        self.assertEqual(names, ["One.docx", "Sample.docx", "Two.docx"])

    def test_cancel_mid_split_leaves_nothing_behind(self):
        info = self.analyze()
        request = {
            "kind": "split", "source": str(self.sample), "fingerprint": info["fingerprint"],
            "total_pages": info["pages"], "page_a": 3, "page_b": 4,
            "output_one": str(self.dir / "One.docx"), "output_two": str(self.dir / "Two.docx"),
        }
        (kind, _data), _progress, job = run(request, cancel_on="progress.part1")
        self.assertEqual(kind, "cancelled")
        self.assertEqual(sorted(p.name for p in self.dir.iterdir()), ["Sample.docx"])
        self.assertFalse(job.job_dir.exists())


if __name__ == "__main__":
    unittest.main()

"""Orchestration tests for split_engine using a simulated Word session.

Microsoft Word cannot run in a test harness, therefore the COM layer is replaced
by a deterministic double that models a document as a character stream with a
fixed number of characters per page. This exercises the parts of the protocol
that carry the real risk: staged commit, verification, rollback, cancellation
and the deletion of the original file.

Run with:

    python tests/test_split_flow.py
"""

from __future__ import annotations

import sys
import tempfile
import threading
import unittest
import zipfile
from contextlib import contextmanager
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC))

import split_engine  # noqa: E402
import utils  # noqa: E402
from split_engine import OperationCancelled, SplitError, SplitRequest  # noqa: E402

CHARS_PER_PAGE = 100
HEADER = "<?xml version='1.0'?><body>"
FOOTER = "</body>"


def write_stream_docx(path: Path, text: str) -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("word/document.xml", HEADER + text + FOOTER)
    return path


def read_stream(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        raw = archive.read("word/document.xml").decode("utf-8")
    return raw[len(HEADER): -len(FOOTER)]


class FakeRange:
    def __init__(self, start: int, end: int) -> None:
        self.Start = start
        self.End = end


class FakeDocument:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.text = read_stream(path)
        self.closed = False

    @property
    def Content(self) -> FakeRange:  # noqa: N802 - mirrors the COM name
        return FakeRange(0, len(self.text))


class FakeSession:
    """Deterministic stand in for word_engine.WordSession."""

    fail_on_open: set[str] = set()
    fail_on_save: set[str] = set()
    fail_on_verify: set[str] = set()
    force_identical = False
    opened: list[str] = []

    def __enter__(self) -> "FakeSession":
        return self

    def __exit__(self, *exc) -> None:
        return None

    @contextmanager
    def open_document(self, path: Path, read_only: bool):
        FakeSession.opened.append(path.name)
        if path.name in FakeSession.fail_on_open:
            from word_engine import DocumentOpenError

            raise DocumentOpenError(f"simulasi gagal membuka {path.name}")
        document = FakeDocument(path)
        try:
            yield document
        finally:
            document.closed = True

    @staticmethod
    def repaginate(document) -> None:
        return None

    def page_count(self, document) -> int:
        if not document.text:
            from word_engine import PaginationError

            raise PaginationError("dokumen kosong")
        return max(1, -(-len(document.text) // CHARS_PER_PAGE))

    def page_start_offset(self, document, page: int) -> int:
        offset = (page - 1) * CHARS_PER_PAGE
        if not 0 < offset < len(document.text):
            from word_engine import PaginationError

            raise PaginationError("offset di luar batas")
        return offset

    def delete_range(self, document, start: int, end: int) -> None:
        if end <= start:
            return
        document.text = document.text[:start] + document.text[end:]

    def save_as_docx(self, document, target: Path) -> None:
        if target.name in FakeSession.fail_on_save:
            from word_engine import SaveError

            raise SaveError(f"simulasi gagal menyimpan {target.name}")
        text = "SAMA" if FakeSession.force_identical else document.text
        write_stream_docx(target, text)

    def verify_document(self, path: Path) -> int:
        if path.name in FakeSession.fail_on_verify:
            from word_engine import DocumentOpenError

            raise DocumentOpenError(f"simulasi gagal verifikasi {path.name}")
        with self.open_document(path, read_only=True) as document:
            return self.page_count(document)


class SplitFlowTestCase(unittest.TestCase):
    def setUp(self) -> None:
        self.dir = Path(tempfile.mkdtemp())
        body = "".join(chr(ord("a") + i) * CHARS_PER_PAGE for i in range(10))
        self.source = write_stream_docx(self.dir / "laporan.docx", body)  # 10 pages
        self.out_one = self.dir / "Part 1.docx"
        self.out_two = self.dir / "Part 2.docx"
        self._original_session = split_engine.WordSession
        split_engine.WordSession = FakeSession
        FakeSession.fail_on_open = set()
        FakeSession.fail_on_save = set()
        FakeSession.fail_on_verify = set()
        FakeSession.force_identical = False
        FakeSession.opened = []

    def tearDown(self) -> None:
        split_engine.WordSession = self._original_session

    def make_request(self, **overrides) -> SplitRequest:
        params = dict(
            source=self.source,
            page_a=5,
            page_b=6,
            output_one=self.out_one,
            output_two=self.out_two,
            delete_original=False,
            overwrite_allowed=False,
        )
        params.update(overrides)
        return SplitRequest(**params)

    def leftovers(self) -> list[str]:
        return sorted(p.name for p in self.dir.iterdir() if p.name.startswith("~"))

    # -- happy paths -------------------------------------------------------

    def test_analysis_reports_page_count(self):
        result = split_engine.analyze_document(self.source)
        self.assertEqual(result.total_pages, 10)

    def test_split_produces_complementary_parts(self):
        result = split_engine.perform_split(self.make_request())
        self.assertTrue(self.out_one.exists() and self.out_two.exists())
        first = read_stream(self.out_one)
        second = read_stream(self.out_two)
        self.assertEqual(len(first), 500)
        self.assertEqual(len(second), 500)
        self.assertEqual(len(first) + len(second), 1000)
        self.assertEqual(result.plan.split_page, 6)
        self.assertEqual(self.leftovers(), [])

    def test_original_is_preserved_by_default(self):
        split_engine.perform_split(self.make_request())
        self.assertTrue(self.source.exists())

    def test_original_deleted_only_when_requested(self):
        result = split_engine.perform_split(self.make_request(delete_original=True))
        self.assertTrue(result.original_deleted)
        self.assertFalse(self.source.exists())
        self.assertTrue(self.out_one.exists() and self.out_two.exists())

    def test_wide_gap_uses_documented_convention(self):
        result = split_engine.perform_split(self.make_request(page_a=4, page_b=6))
        self.assertEqual(result.plan.split_page, 6)

    def test_reversed_pages_are_normalised(self):
        result = split_engine.perform_split(self.make_request(page_a=6, page_b=5))
        self.assertEqual(result.plan.split_page, 6)

    # -- failure paths -----------------------------------------------------

    def test_missing_source_is_reported(self):
        self.source.unlink()
        with self.assertRaises(SplitError):
            split_engine.perform_split(self.make_request())

    def test_corrupt_source_is_reported(self):
        self.source.write_bytes(b"not a zip")
        with self.assertRaises(SplitError):
            split_engine.perform_split(self.make_request())

    def test_existing_output_blocks_without_permission(self):
        write_stream_docx(self.out_one, "lama")
        with self.assertRaises(SplitError):
            split_engine.perform_split(self.make_request())
        self.assertEqual(read_stream(self.out_one), "lama")

    def test_existing_output_replaced_when_allowed(self):
        write_stream_docx(self.out_one, "lama")
        split_engine.perform_split(self.make_request(overwrite_allowed=True))
        self.assertNotEqual(read_stream(self.out_one), "lama")

    def test_failure_while_building_part_two_leaves_no_output(self):
        # SaveAs receives the temporary name, so the double targets that.
        original_save = FakeSession.save_as_docx

        def failing_save(self_, document, target: Path) -> None:
            if "part2" in target.name:
                from word_engine import SaveError

                raise SaveError("simulasi kegagalan pada Part 2")
            original_save(self_, document, target)

        FakeSession.save_as_docx = failing_save
        try:
            with self.assertRaises(SplitError):
                split_engine.perform_split(self.make_request())
        finally:
            FakeSession.save_as_docx = original_save

        self.assertFalse(self.out_one.exists())
        self.assertFalse(self.out_two.exists())
        self.assertTrue(self.source.exists())
        self.assertEqual(self.leftovers(), [])

    def test_verification_failure_aborts_and_cleans_up(self):
        original_verify = FakeSession.verify_document

        def failing_verify(self_, path: Path) -> int:
            from word_engine import DocumentOpenError

            raise DocumentOpenError("simulasi gagal verifikasi")

        FakeSession.verify_document = failing_verify
        try:
            with self.assertRaises(SplitError):
                split_engine.perform_split(self.make_request(delete_original=True))
        finally:
            FakeSession.verify_document = original_verify

        self.assertTrue(self.source.exists(), "file asli harus tetap ada")
        self.assertFalse(self.out_one.exists())
        self.assertFalse(self.out_two.exists())
        self.assertEqual(self.leftovers(), [])

    def test_silent_delete_failure_is_rejected(self):
        """If the deletion step does nothing, both parts keep the full document."""
        original_delete = FakeSession.delete_range

        def noop_delete(self_, document, start, end):
            return None

        FakeSession.delete_range = noop_delete
        try:
            with self.assertRaises(SplitError):
                split_engine.perform_split(self.make_request())
        finally:
            FakeSession.delete_range = original_delete
        self.assertFalse(self.out_one.exists())
        self.assertFalse(self.out_two.exists())
        self.assertTrue(self.source.exists())
        self.assertEqual(self.leftovers(), [])

    def test_byte_identical_outputs_only_warn(self):
        FakeSession.force_identical = True
        result = split_engine.perform_split(self.make_request())
        self.assertTrue(result.warnings)
        self.assertTrue(self.out_one.exists() and self.out_two.exists())

    def test_out_of_range_pages_are_rejected(self):
        from validation import ValidationError

        with self.assertRaises((SplitError, ValidationError)):
            split_engine.perform_split(self.make_request(page_a=20, page_b=21))
        self.assertTrue(self.source.exists())

    def test_cancellation_stops_before_commit(self):
        event = threading.Event()
        event.set()
        with self.assertRaises(OperationCancelled):
            split_engine.perform_split(self.make_request(), cancel_event=event)
        self.assertFalse(self.out_one.exists())
        self.assertFalse(self.out_two.exists())
        self.assertTrue(self.source.exists())
        self.assertEqual(self.leftovers(), [])

    def test_progress_messages_are_emitted(self):
        messages: list[str] = []
        split_engine.perform_split(self.make_request(), progress=messages.append)
        self.assertTrue(any("Part 1" in m for m in messages))
        self.assertTrue(any("Part 2" in m for m in messages))

    def test_temp_artifacts_use_the_destination_volume(self):
        target = self.dir / "Part 1.docx"
        temp = utils.temp_sibling_path(target, "part1")
        self.assertEqual(temp.parent, target.parent)


if __name__ == "__main__":
    unittest.main(verbosity=2)

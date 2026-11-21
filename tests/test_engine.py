"""Orchestration tests for split_engine with a simulated Word session.

Word cannot run in every test environment, so the session is replaced by a
deterministic double that models a document as a character stream with 100
characters per page. This exercises the parts of the protocol that carry the
real risk: snapshotting, fingerprints, transactional commit with backups,
rollback, cancellation and removal of the original.
"""

from __future__ import annotations

import sys
import tempfile
import threading
import unittest
import zipfile
from pathlib import Path
from unittest import mock

SRC = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC))

import split_engine  # noqa: E402
import utils  # noqa: E402
import winsys  # noqa: E402
from errors import AppError, OperationCancelled  # noqa: E402
from split_engine import SplitRequest  # noqa: E402

PAGE = 100
HEADER, FOOTER = "<?xml version='1.0'?><body>", "</body>"


def write_stream(path: Path, text: str) -> Path:
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("word/document.xml", HEADER + text + FOOTER)
    return path


def read_stream(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        raw = archive.read("word/document.xml").decode("utf-8")
    return raw[len(HEADER): -len(FOOTER)]


class FakeDocument:
    def __init__(self, path: Path) -> None:
        self.text = read_stream(path)


class Report:
    def __init__(self, unlinked: int = 0) -> None:
        self.fields_unlinked = unlinked
        self.steps = []


class FakeSession:
    fail_save: set = set()
    fail_reopen: set = set()
    silent_cut = False
    page_overflow = 0
    unlinked = 0
    opened: list = []

    def __init__(self, job=None) -> None:
        self.job = job

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return None

    def open(self, path: Path):
        FakeSession.opened.append(path)
        return FakeDocument(path)

    def close_document(self, document) -> None:
        pass

    def page_count(self, document) -> int:
        return max(1, -(-len(document.text) // PAGE)) + (FakeSession.page_overflow if len(document.text) < 2000 else 0)

    def page_start(self, document, page: int) -> int:
        offset = (page - 1) * PAGE
        if not 0 < offset < len(document.text):
            raise AppError("err.empty_part")
        return offset

    def cut_keep_before(self, document, split: int):
        if not FakeSession.silent_cut:
            document.text = document.text[:split]
        return Report(FakeSession.unlinked)

    def cut_keep_from(self, document, split: int):
        if not FakeSession.silent_cut:
            document.text = document.text[split:]
        return Report()

    def save_docx(self, document, target: Path) -> None:
        if any(tag in target.name for tag in FakeSession.fail_save):
            raise AppError("err.save_failed")
        write_stream(target, document.text)

    def reopen_check(self, path: Path, label: str) -> None:
        if any(tag in path.name for tag in FakeSession.fail_reopen):
            raise AppError("err.verify_failed", label=label)


def page_text(pages: int) -> str:
    return "".join(chr(65 + (i % 26)) * PAGE for i in range(pages))


class EngineTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.job_dir = self.dir / "job"
        self.source = write_stream(self.dir / "source.docx", page_text(10))
        FakeSession.fail_save = set()
        FakeSession.fail_reopen = set()
        FakeSession.silent_cut = False
        FakeSession.page_overflow = 0
        FakeSession.unlinked = 0
        FakeSession.opened = []
        # Never touch the real Recycle Bin from tests.
        patcher = mock.patch.object(winsys, "drive_has_recycle_bin", return_value=False)
        patcher.start()
        self.addCleanup(patcher.stop)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def analyze(self):
        return split_engine.analyze(self.source, self.job_dir, session_factory=FakeSession)

    def request(self, a=5, b=6, **overrides) -> SplitRequest:
        info = self.analyze()
        values = dict(
            source=self.source, job_dir=self.job_dir,
            fingerprint=utils.Fingerprint.from_list(info["fingerprint"]), total_pages=info["pages"],
            page_a=a, page_b=b, output_one=self.dir / "Part 1.docx", output_two=self.dir / "Part 2.docx",
        )
        values.update(overrides)
        return SplitRequest(**values)

    def finals_exist(self) -> tuple:
        return (self.dir / "Part 1.docx").exists(), (self.dir / "Part 2.docx").exists()

    # -- analysis -----------------------------------------------------------

    def test_analysis_reports_pages_and_never_opens_the_original(self):
        info = self.analyze()
        self.assertEqual(info["pages"], 10)
        self.assertEqual(len(info["fingerprint"]), 2)
        self.assertTrue(all(path != self.source for path in FakeSession.opened))
        self.assertTrue(all(path.parent == self.job_dir for path in FakeSession.opened))

    def test_analysis_rejects_damaged_package_before_word(self):
        self.source.write_bytes(b"not a zip at all")
        with self.assertRaises(AppError) as ctx:
            self.analyze()
        self.assertEqual(ctx.exception.key, "err.not_zip")
        self.assertEqual(FakeSession.opened, [])

    # -- split ----------------------------------------------------------------

    def test_split_parts_complement_each_other(self):
        result = split_engine.split(self.request(5, 6), session_factory=FakeSession)
        one, two = read_stream(self.dir / "Part 1.docx"), read_stream(self.dir / "Part 2.docx")
        self.assertEqual(one + two, page_text(10))
        self.assertEqual((result.pages_one, result.pages_two), (5, 5))
        self.assertEqual(result.plan.split_page, 6)
        self.assertEqual(result.warnings, [])
        self.assertTrue(self.source.exists())
        self.assertEqual(result.original_removed, "kept")

    def test_reversed_pages_give_the_same_result(self):
        result = split_engine.split(self.request(6, 5), session_factory=FakeSession)
        self.assertEqual(result.plan.split_page, 6)

    def test_gap_keeps_midpoint_in_part_one(self):
        result = split_engine.split(self.request(4, 6), session_factory=FakeSession)
        self.assertEqual(result.pages_one, 5)

    def test_no_staging_files_left_in_destination(self):
        split_engine.split(self.request(), session_factory=FakeSession)
        names = sorted(p.name for p in self.dir.iterdir() if p.is_file())
        self.assertEqual(names, ["Part 1.docx", "Part 2.docx", "source.docx"])

    def test_source_changed_after_analysis(self):
        request = self.request()
        write_stream(self.source, page_text(12))
        with self.assertRaises(AppError) as ctx:
            split_engine.split(request, session_factory=FakeSession)
        self.assertEqual(ctx.exception.key, "err.source_changed")
        self.assertEqual(self.finals_exist(), (False, False))

    def test_page_count_drift_is_treated_as_change(self):
        request = self.request()
        request.total_pages = 11
        with self.assertRaises(AppError) as ctx:
            split_engine.split(request, session_factory=FakeSession)
        self.assertEqual(ctx.exception.key, "err.source_changed")

    def test_existing_outputs_refused_without_overwrite(self):
        (self.dir / "Part 1.docx").write_bytes(b"user data")
        with self.assertRaises(AppError) as ctx:
            split_engine.split(self.request(), session_factory=FakeSession)
        self.assertEqual(ctx.exception.key, "err.outputs_exist")
        self.assertEqual((self.dir / "Part 1.docx").read_bytes(), b"user data")

    def test_existing_outputs_replaced_with_overwrite(self):
        (self.dir / "Part 1.docx").write_bytes(b"user data")
        split_engine.split(self.request(overwrite_allowed=True), session_factory=FakeSession)
        self.assertEqual(read_stream(self.dir / "Part 1.docx"), page_text(5))
        self.assertFalse(any(p.name.startswith("~wordsplitter") for p in self.dir.iterdir()))

    def test_failure_building_part_two_writes_nothing(self):
        FakeSession.fail_save = {"part2"}
        with self.assertRaises(AppError) as ctx:
            split_engine.split(self.request(), session_factory=FakeSession)
        self.assertEqual(ctx.exception.key, "err.save_failed")
        self.assertEqual(self.finals_exist(), (False, False))
        self.assertTrue(self.source.exists())

    def test_failed_verification_keeps_original_even_with_delete(self):
        FakeSession.fail_reopen = {"part1"}
        with self.assertRaises(AppError) as ctx:
            split_engine.split(self.request(delete_original=True), session_factory=FakeSession)
        self.assertEqual(ctx.exception.key, "err.verify_failed")
        self.assertTrue(self.source.exists())
        self.assertEqual(self.finals_exist(), (False, False))

    def test_silent_cut_failure_is_detected(self):
        FakeSession.silent_cut = True
        with self.assertRaises(AppError) as ctx:
            split_engine.split(self.request(), session_factory=FakeSession)
        self.assertEqual(ctx.exception.key, "err.both_full")
        self.assertEqual(self.finals_exist(), (False, False))

    def test_commit_failure_rolls_back_and_restores_previous_files(self):
        (self.dir / "Part 1.docx").write_bytes(b"old one")
        (self.dir / "Part 2.docx").write_bytes(b"old two")
        real_copy = utils.copy_into_place
        calls = {"n": 0}

        def flaky(source, final, overwrite):
            calls["n"] += 1
            if calls["n"] == 2:
                raise PermissionError("locked by another program")
            return real_copy(source, final, overwrite)

        with mock.patch.object(utils, "copy_into_place", side_effect=flaky):
            with self.assertRaises(AppError) as ctx:
                split_engine.split(self.request(overwrite_allowed=True), session_factory=FakeSession)
        self.assertEqual(ctx.exception.key, "err.commit_failed")
        self.assertEqual((self.dir / "Part 1.docx").read_bytes(), b"old one")
        self.assertEqual((self.dir / "Part 2.docx").read_bytes(), b"old two")
        self.assertFalse(any(p.name.startswith("~wordsplitter") for p in self.dir.iterdir()))

    def test_delete_original_permanently_on_drive_without_bin(self):
        result = split_engine.split(self.request(delete_original=True), session_factory=FakeSession)
        self.assertEqual(result.original_removed, "deleted")
        self.assertFalse(self.source.exists())

    def test_delete_original_uses_recycle_bin_when_available(self):
        def fake_recycle(path):
            Path(path).unlink()
            return True

        with mock.patch.object(winsys, "drive_has_recycle_bin", return_value=True), \
                mock.patch.object(winsys, "move_to_recycle_bin", side_effect=fake_recycle) as recycle:
            result = split_engine.split(self.request(delete_original=True), session_factory=FakeSession)
        self.assertEqual(result.original_removed, "recycled")
        recycle.assert_called_once()

    def test_recycle_failure_is_a_warning_not_an_error(self):
        with mock.patch.object(winsys, "drive_has_recycle_bin", return_value=True), \
                mock.patch.object(winsys, "move_to_recycle_bin", return_value=False):
            result = split_engine.split(self.request(delete_original=True), session_factory=FakeSession)
        self.assertEqual(result.original_removed, "failed")
        self.assertIn("warn.delete_failed", [w["key"] for w in result.warnings])
        self.assertTrue(self.source.exists())
        self.assertEqual(self.finals_exist(), (True, True))

    def test_layout_drift_and_unlinked_fields_become_warnings(self):
        FakeSession.page_overflow = 1
        FakeSession.unlinked = 1
        result = split_engine.split(self.request(), session_factory=FakeSession)
        keys = [w["key"] for w in result.warnings]
        self.assertIn("warn.pages_differ", keys)
        self.assertIn("warn.fields_unlinked", keys)

    def test_out_of_range_rejected_before_any_output(self):
        with self.assertRaises(AppError) as ctx:
            split_engine.split(self.request(9, 40), session_factory=FakeSession)
        self.assertEqual(ctx.exception.key, "err.page_out_of_range")
        self.assertEqual(self.finals_exist(), (False, False))

    def test_cancel_before_word_starts(self):
        cancel = threading.Event()
        cancel.set()
        with self.assertRaises(OperationCancelled):
            split_engine.split(self.request(), session_factory=FakeSession, cancel=cancel)
        self.assertEqual(self.finals_exist(), (False, False))

    def test_progress_and_commit_events(self):
        events = []
        split_engine.split(self.request(), emit=lambda kind, **data: events.append((kind, data)),
                           session_factory=FakeSession)
        keys = [d.get("key") for kind, d in events if kind == "progress"]
        for key in ("progress.part1", "progress.part2", "progress.verifying", "progress.saving"):
            self.assertIn(key, keys)
        kinds = [kind for kind, _ in events]
        self.assertIn("commit", kinds)
        self.assertLess(kinds.index("commit"), len(kinds))
        percentages = [d["pct"] for kind, d in events if kind == "progress"]
        self.assertEqual(percentages, sorted(percentages))

    def test_request_round_trip(self):
        request = self.request()
        data = {
            "source": str(request.source), "job_dir": str(request.job_dir),
            "fingerprint": request.fingerprint.as_list(), "total_pages": 10, "page_a": 5, "page_b": 6,
            "output_one": str(request.output_one), "output_two": str(request.output_two),
        }
        again = SplitRequest.from_dict(data)
        self.assertEqual(again.fingerprint, request.fingerprint)
        self.assertFalse(again.delete_original)


if __name__ == "__main__":
    unittest.main()

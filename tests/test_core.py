"""Pure logic tests: validation rules, split arithmetic, package checks, helpers.

Run with:  py -3 -m unittest discover -s tests -v
None of these tests needs Microsoft Word.
"""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC))

import utils  # noqa: E402
import validation  # noqa: E402
from errors import AppError, ValidationError  # noqa: E402

MINIMAL_DOCUMENT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    "<w:body><w:p><w:r><w:t>Halo</w:t></w:r></w:p></w:body></w:document>"
)
MINIMAL_CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/></Types>'
)


def make_docx(path: Path, payload: str = MINIMAL_DOCUMENT_XML) -> Path:
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", MINIMAL_CONTENT_TYPES)
        archive.writestr("word/document.xml", payload)
    return path


class TempDirTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()


class TestPageParsing(unittest.TestCase):
    def test_accepts_plain_integer(self):
        self.assertEqual(validation.parse_page_number(" 12 ", "label.page_a"), 12)

    def test_accepts_digits_typed_with_other_keyboards(self):
        for raw, expected in (("١٢", 12), ("۱۲", 12), ("१२", 12), ("１２", 12)):
            with self.subTest(raw=raw):
                self.assertEqual(validation.parse_page_number(raw, "label.page_a"), expected)

    def test_rejects_empty(self):
        with self.assertRaises(ValidationError) as ctx:
            validation.parse_page_number("  ", "label.page_a")
        self.assertEqual(ctx.exception.key, "err.page_empty")
        self.assertEqual(ctx.exception.params["label"], "label.page_a")

    def test_rejects_non_numeric(self):
        for value in ("abc", "3.5", "1e3", "10a", "--4", "1 2"):
            with self.subTest(value=value):
                with self.assertRaises(ValidationError) as ctx:
                    validation.parse_page_number(value, "label.page_a")
                self.assertEqual(ctx.exception.key, "err.page_not_int")

    def test_rejects_zero_and_negative(self):
        for value in ("0", "-1", "-25"):
            with self.subTest(value=value):
                with self.assertRaises(ValidationError) as ctx:
                    validation.parse_page_number(value, "label.page_b")
                self.assertEqual(ctx.exception.key, "err.page_zero")


class TestSplitPlan(unittest.TestCase):
    def test_adjacent_pages(self):
        plan = validation.compute_split_plan(10, 11, 40)
        self.assertEqual(plan.split_page, 11)
        self.assertEqual(plan.part_one, (1, 10))
        self.assertEqual(plan.part_two, (11, 40))

    def test_gap_of_two_keeps_midpoint_in_part_one(self):
        self.assertEqual(validation.compute_split_plan(10, 12, 40).split_page, 12)

    def test_gap_of_three(self):
        self.assertEqual(validation.compute_split_plan(10, 13, 40).split_page, 12)

    def test_order_is_normalised(self):
        self.assertEqual(validation.compute_split_plan(11, 10, 40).split_page,
                         validation.compute_split_plan(10, 11, 40).split_page)

    def test_first_and_last_boundary(self):
        self.assertEqual(validation.compute_split_plan(1, 2, 2).split_page, 2)
        self.assertEqual(validation.compute_split_plan(39, 40, 40).split_page, 40)

    def test_identical_pages_rejected(self):
        with self.assertRaises(ValidationError) as ctx:
            validation.compute_split_plan(5, 5, 10)
        self.assertEqual(ctx.exception.key, "err.page_same")

    def test_out_of_range_names_the_field(self):
        with self.assertRaises(ValidationError) as ctx:
            validation.compute_split_plan(5, 50, 40)
        self.assertEqual(ctx.exception.key, "err.page_out_of_range")
        self.assertEqual(ctx.exception.params, {"label": "label.page_b", "value": 50, "total": 40})

    def test_single_page_document(self):
        with self.assertRaises(ValidationError) as ctx:
            validation.compute_split_plan(1, 2, 1)
        self.assertEqual(ctx.exception.key, "err.single_page")

    def test_exhaustive_no_empty_part(self):
        for total in range(2, 40):
            for a in range(1, total + 1):
                for b in range(1, total + 1):
                    if a == b:
                        continue
                    plan = validation.compute_split_plan(a, b, total)
                    self.assertGreaterEqual(plan.split_page - 1, 1)
                    self.assertGreaterEqual(total - plan.split_page + 1, 1)
                    low, high = sorted((a, b))
                    self.assertLessEqual(low, plan.split_page - 1)
                    self.assertGreaterEqual(high, plan.split_page)


class TestOutputNames(unittest.TestCase):
    def test_extension_added_and_normalised(self):
        self.assertEqual(validation.validate_output_name("Part 1", "label.name1"), "Part 1.docx")
        self.assertEqual(validation.validate_output_name("Part 1.DOCX", "label.name1"), "Part 1.docx")
        self.assertEqual(validation.validate_output_name("a.b.docx", "label.name1"), "a.b.docx")

    def test_unicode_names_allowed(self):
        for name in ("Bab 1–3", "第 1 部分", "الجزء 1", "भाग 1", "Część 1", "ส่วนที่ 1"):
            with self.subTest(name=name):
                self.assertTrue(validation.validate_output_name(name, "label.name1").endswith(".docx"))

    def test_invalid_characters(self):
        for char in '\\/:*?"<>|':
            with self.subTest(char=char):
                with self.assertRaises(ValidationError) as ctx:
                    validation.validate_output_name(f"bad{char}name", "label.name1")
                self.assertEqual(ctx.exception.key, "err.name_invalid_chars")

    def test_rules(self):
        cases = {
            "": "err.name_empty",
            "   ": "err.name_empty",
            ".docx": "err.name_only_ext",
            "Part 1.": "err.name_trailing",
            "Part 1 .docx": "err.name_trailing",
            "CON": "err.name_reserved",
            "nul.docx": "err.name_reserved",
            "COM1": "err.name_reserved",
            "LPT9": "err.name_reserved",
            "con.backup": "err.name_reserved",
            "a\x07b": "err.name_control",
            "x" * 181: "err.name_too_long",
        }
        for raw, key in cases.items():
            with self.subTest(raw=raw[:20]):
                with self.assertRaises(ValidationError) as ctx:
                    validation.validate_output_name(raw, "label.name1")
                self.assertEqual(ctx.exception.key, key)


class TestOutputPaths(TempDirTest):
    def test_resolves_inside_chosen_directory(self):
        source = self.dir / "src.docx"
        out = self.dir / "out"
        out.mkdir()
        one, two = validation.resolve_output_paths(source, out, "A", "B")
        self.assertEqual(one, out / "A.docx")
        self.assertEqual(two, out / "B.docx")

    def test_case_insensitive_collision(self):
        with self.assertRaises(ValidationError) as ctx:
            validation.resolve_output_paths(self.dir / "s.docx", self.dir, "Part", "PART")
        self.assertEqual(ctx.exception.key, "err.names_same")

    def test_overwriting_source_rejected(self):
        source = self.dir / "Report.docx"
        with self.assertRaises(ValidationError) as ctx:
            validation.resolve_output_paths(source, self.dir, "report", "other")
        self.assertEqual(ctx.exception.key, "err.name_is_source")

    def test_path_too_long_for_word(self):
        deep = Path("C:/") / ("d" * 120) / ("e" * 120)
        with self.assertRaises(ValidationError) as ctx:
            validation.resolve_output_paths(deep / "s.docx", deep, "x" * 20, "y")
        self.assertEqual(ctx.exception.key, "err.output_path_too_long")

    def test_existing_outputs(self):
        (self.dir / "A.docx").write_bytes(b"x")
        found = validation.existing_outputs(self.dir / "A.docx", self.dir / "B.docx")
        self.assertEqual(found, [self.dir / "A.docx"])


class TestInputValidation(TempDirTest):
    def test_valid_docx(self):
        path = make_docx(self.dir / "ok.docx")
        self.assertEqual(validation.validate_input_file(str(path)), path.resolve())

    def expect(self, path: Path, key: str) -> None:
        with self.assertRaises(ValidationError) as ctx:
            validation.validate_input_file(str(path))
        self.assertEqual(ctx.exception.key, key)

    def test_missing_and_wrong_extension(self):
        self.expect(self.dir / "missing.docx", "err.not_found")
        other = self.dir / "file.doc"
        other.write_bytes(b"x")
        self.expect(other, "err.not_docx_ext")

    def test_empty_file(self):
        path = self.dir / "empty.docx"
        path.write_bytes(b"")
        self.expect(path, "err.empty_file")

    def test_not_a_zip(self):
        path = self.dir / "text.docx"
        path.write_bytes(b"this is plain text pretending to be docx")
        self.expect(path, "err.not_zip")

    def test_encrypted_or_legacy_compound_file(self):
        path = self.dir / "locked.docx"
        path.write_bytes(utils.CFB_SIGNATURE + b"\x00" * 600)
        self.expect(path, "err.encrypted_or_legacy")

    def test_zip_without_document(self):
        path = self.dir / "nodoc.docx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("[Content_Types].xml", MINIMAL_CONTENT_TYPES)
        self.expect(path, "err.zip_missing")

    def test_truncated_archive(self):
        good = make_docx(self.dir / "good.docx", MINIMAL_DOCUMENT_XML * 50)
        data = good.read_bytes()
        broken = self.dir / "broken.docx"
        broken.write_bytes(data[: len(data) // 2])
        self.expect(broken, "err.zip_corrupt")

    def test_corrupted_entry_crc(self):
        good = make_docx(self.dir / "crc.docx", MINIMAL_DOCUMENT_XML * 40)
        data = bytearray(good.read_bytes())
        offset = data.find(b"word/document.xml") + len("word/document.xml") + 12
        data[offset] ^= 0xFF
        bad = self.dir / "crc_bad.docx"
        bad.write_bytes(bytes(data))
        self.expect(bad, "err.zip_corrupt")

    def test_no_selection(self):
        with self.assertRaises(ValidationError) as ctx:
            validation.validate_input_file("")
        self.assertEqual(ctx.exception.key, "err.no_file")


class TestUtils(TempDirTest):
    def test_format_size(self):
        self.assertEqual(utils.format_size(512), "512 B")
        self.assertEqual(utils.format_size(1536), "1.50 KB")
        self.assertEqual(utils.format_size(1536, ("o", "Ko", "Mo", "Go", "To"), ","), "1,50 Ko")
        self.assertEqual(utils.format_size(-1), "?")

    def test_extension_helpers(self):
        self.assertEqual(utils.ensure_docx_extension(" a "), "a.docx")
        self.assertEqual(utils.strip_docx_extension("a.DocX"), "a")

    def test_unique_names_never_repeat(self):
        names = {utils.unique_name("part1") for _ in range(500)}
        self.assertEqual(len(names), 500)

    def test_temp_sibling_on_same_volume(self):
        target = self.dir / "final.docx"
        staging = utils.temp_sibling_path(target)
        self.assertEqual(staging.parent, target.parent)
        self.assertNotEqual(staging, target)

    def test_copy_into_place_is_complete_and_leaves_no_staging(self):
        source = self.dir / "src.bin"
        source.write_bytes(os.urandom(3 * 1024 * 1024 + 7))
        final = self.dir / "final.bin"
        utils.copy_into_place(source, final, overwrite=False)
        self.assertEqual(final.read_bytes(), source.read_bytes())
        self.assertEqual(sorted(p.name for p in self.dir.iterdir()), ["final.bin", "src.bin"])

    def test_copy_into_place_refuses_existing_without_overwrite(self):
        source = self.dir / "src.bin"
        source.write_bytes(b"new")
        final = self.dir / "final.bin"
        final.write_bytes(b"old")
        with self.assertRaises(FileExistsError):
            utils.copy_into_place(source, final, overwrite=False)
        self.assertEqual(final.read_bytes(), b"old")
        self.assertEqual(len(list(self.dir.iterdir())), 2)

    def test_safe_unlink_and_remove_tree_never_raise(self):
        utils.safe_unlink(None)
        utils.safe_unlink(self.dir / "nope")
        self.assertTrue(utils.remove_tree(self.dir / "nope"))
        nested = self.dir / "job" / "deep"
        nested.mkdir(parents=True)
        (nested / "f").write_bytes(b"x")
        self.assertTrue(utils.remove_tree(self.dir / "job"))

    def test_owner_file_variants(self):
        names = {p.name for p in utils.word_owner_files(self.dir / "Report.docx")}
        self.assertEqual(names, {"~$Report.docx", "~$port.docx"})

    def test_identical_detection(self):
        a, b, c = (self.dir / n for n in ("a", "b", "c"))
        a.write_bytes(b"same")
        b.write_bytes(b"same")
        c.write_bytes(b"diff")
        self.assertTrue(utils.files_are_identical(a, b))
        self.assertFalse(utils.files_are_identical(a, c))
        self.assertFalse(utils.files_are_identical(a, self.dir / "missing"))

    def test_writable_directory_check(self):
        utils.check_writable_directory(self.dir)
        self.assertEqual(list(self.dir.iterdir()), [])
        with self.assertRaises(AppError) as ctx:
            utils.check_writable_directory(self.dir / "missing")
        self.assertEqual(ctx.exception.key, "err.dest_missing")

    def test_fingerprint_detects_change(self):
        path = self.dir / "f.docx"
        path.write_bytes(b"one")
        first = utils.Fingerprint.of(path)
        self.assertEqual(utils.Fingerprint.from_list(first.as_list()), first)
        path.write_bytes(b"one two")
        self.assertNotEqual(utils.Fingerprint.of(path), first)

    def test_space_estimate_is_conservative(self):
        self.assertGreaterEqual(validation.estimate_required_space(10 * 1024 * 1024), 5 * 10 * 1024 * 1024)


class TestErrors(unittest.TestCase):
    def test_round_trip(self):
        error = AppError("err.disk_space", detail="x", needed_bytes=10, free_bytes=2)
        again = AppError.from_dict(error.to_dict())
        self.assertEqual((again.key, again.params, again.detail), (error.key, error.params, error.detail))

    def test_cancellation_is_not_swallowed_by_except_exception(self):
        from errors import OperationCancelled

        self.assertFalse(issubclass(OperationCancelled, Exception))


if __name__ == "__main__":
    unittest.main()

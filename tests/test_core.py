"""Platform independent self tests for the pure logic of WordSplitter.

These tests exercise everything that does not require Microsoft Word: file name
validation, page parsing, split point arithmetic, .docx container inspection,
temporary path allocation and byte comparison. Run with:

    python tests/test_core.py

The Word dependent behaviour is covered by the manual test plan in TEST_PLAN.md.
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
from validation import ValidationError  # noqa: E402

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
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", MINIMAL_CONTENT_TYPES)
        archive.writestr("word/document.xml", payload)
    return path


class TestPageParsing(unittest.TestCase):
    def test_accepts_plain_integer(self):
        self.assertEqual(validation.parse_page_number(" 12 ", "Page A"), 12)

    def test_rejects_empty(self):
        with self.assertRaises(ValidationError):
            validation.parse_page_number("", "Page A")

    def test_rejects_non_numeric(self):
        for value in ("abc", "3.5", "1e3", "10a", "--4"):
            with self.subTest(value=value):
                with self.assertRaises(ValidationError):
                    validation.parse_page_number(value, "Page A")

    def test_rejects_zero_and_negative(self):
        for value in ("0", "-1", "-25"):
            with self.subTest(value=value):
                with self.assertRaises(ValidationError):
                    validation.parse_page_number(value, "Page B")


class TestSplitPlan(unittest.TestCase):
    def test_adjacent_pages(self):
        plan = validation.compute_split_plan(10, 11, 40)
        self.assertEqual(plan.split_page, 11)

    def test_gap_of_two_keeps_midpoint_in_part_one(self):
        plan = validation.compute_split_plan(10, 12, 40)
        self.assertEqual(plan.split_page, 12)

    def test_order_is_normalised(self):
        self.assertEqual(
            validation.compute_split_plan(11, 10, 40).split_page,
            validation.compute_split_plan(10, 11, 40).split_page,
        )

    def test_first_boundary(self):
        self.assertEqual(validation.compute_split_plan(1, 2, 5).split_page, 2)

    def test_last_boundary(self):
        self.assertEqual(validation.compute_split_plan(9, 10, 10).split_page, 10)

    def test_equal_pages_rejected(self):
        with self.assertRaises(ValidationError):
            validation.compute_split_plan(7, 7, 20)

    def test_out_of_range_rejected(self):
        with self.assertRaises(ValidationError):
            validation.compute_split_plan(19, 21, 20)

    def test_single_page_document_rejected(self):
        with self.assertRaises(ValidationError):
            validation.compute_split_plan(1, 2, 1)

    def test_split_page_never_leaves_empty_part(self):
        for total in range(2, 30):
            for a in range(1, total + 1):
                for b in range(1, total + 1):
                    if a == b:
                        continue
                    try:
                        plan = validation.compute_split_plan(a, b, total)
                    except ValidationError:
                        continue
                    self.assertGreaterEqual(plan.split_page, 2)
                    self.assertLessEqual(plan.split_page, total)


class TestOutputNames(unittest.TestCase):
    def test_extension_added(self):
        self.assertEqual(validation.validate_output_name("Part 1", "File 1 Name"), "Part 1.docx")

    def test_extension_not_duplicated(self):
        self.assertEqual(
            validation.validate_output_name("Part 1.DOCX", "File 1 Name"), "Part 1.docx"
        )

    def test_illegal_characters_rejected(self):
        for value in ('a/b', 'a\\b', 'a:b', 'a*b', 'a?b', 'a"b', 'a<b', 'a>b', 'a|b'):
            with self.subTest(value=value):
                with self.assertRaises(ValidationError):
                    validation.validate_output_name(value, "File 1 Name")

    def test_empty_rejected(self):
        for value in ("", "   ", ".docx"):
            with self.subTest(value=value):
                with self.assertRaises(ValidationError):
                    validation.validate_output_name(value, "File 1 Name")

    def test_reserved_names_rejected(self):
        for value in ("CON", "nul", "COM1", "LPT9"):
            with self.subTest(value=value):
                with self.assertRaises(ValidationError):
                    validation.validate_output_name(value, "File 1 Name")

    def test_trailing_dot_or_space_rejected(self):
        for value in ("Part 1.", "Part 1 "):
            with self.subTest(value=value):
                with self.assertRaises(ValidationError):
                    validation.validate_output_name(value + ".docx", "File 1 Name")

    def test_control_characters_rejected(self):
        with self.assertRaises(ValidationError):
            validation.validate_output_name("Part\x001", "File 1 Name")

    def test_overlong_rejected(self):
        with self.assertRaises(ValidationError):
            validation.validate_output_name("x" * 500, "File 1 Name")


class TestOutputPaths(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())
        self.source = make_docx(self.dir / "laporan.docx")

    def test_paths_resolve_next_to_source(self):
        one, two = validation.resolve_output_paths(self.source, "Part 1", "Part 2")
        self.assertEqual(one.parent, self.dir)
        self.assertEqual(two.parent, self.dir)
        self.assertEqual(one.name, "Part 1.docx")

    def test_identical_names_rejected(self):
        with self.assertRaises(ValidationError):
            validation.resolve_output_paths(self.source, "Part 1", "part 1")

    def test_overwriting_source_rejected(self):
        with self.assertRaises(ValidationError):
            validation.resolve_output_paths(self.source, "laporan", "Part 2")

    def test_existing_outputs_detected(self):
        one, two = validation.resolve_output_paths(self.source, "Part 1", "Part 2")
        make_docx(one)
        self.assertEqual(validation.existing_outputs(one, two), [one])


class TestInputValidation(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())

    def test_valid_docx_accepted(self):
        path = make_docx(self.dir / "ok.docx")
        self.assertEqual(validation.validate_input_file(str(path)).name, "ok.docx")

    def test_missing_file_rejected(self):
        with self.assertRaises(ValidationError):
            validation.validate_input_file(str(self.dir / "hilang.docx"))

    def test_wrong_extension_rejected(self):
        path = self.dir / "berkas.doc"
        path.write_bytes(b"anything")
        with self.assertRaises(ValidationError):
            validation.validate_input_file(str(path))

    def test_corrupt_container_rejected(self):
        path = self.dir / "rusak.docx"
        path.write_bytes(b"this is definitely not a zip archive")
        with self.assertRaises(ValidationError):
            validation.validate_input_file(str(path))

    def test_zero_byte_rejected(self):
        path = self.dir / "kosong.docx"
        path.write_bytes(b"")
        with self.assertRaises(ValidationError):
            validation.validate_input_file(str(path))

    def test_zip_without_document_xml_rejected(self):
        path = self.dir / "palsu.docx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("hello.txt", "hi")
        with self.assertRaises(ValidationError):
            validation.validate_input_file(str(path))

    def test_empty_selection_rejected(self):
        with self.assertRaises(ValidationError):
            validation.validate_input_file("")


class TestUtils(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp())

    def test_human_readable_size(self):
        self.assertEqual(utils.human_readable_size(0), "0 B")
        self.assertEqual(utils.human_readable_size(1023), "1023 B")
        self.assertEqual(utils.human_readable_size(1024), "1.00 KB")
        self.assertTrue(utils.human_readable_size(5 * 1024 * 1024).endswith("MB"))

    def test_ensure_extension(self):
        self.assertEqual(utils.ensure_docx_extension("a"), "a.docx")
        self.assertEqual(utils.ensure_docx_extension("a.docx"), "a.docx")
        self.assertEqual(utils.ensure_docx_extension(" a.DocX "), "a.docx")

    def test_temp_sibling_is_unique_and_local(self):
        target = self.dir / "Part 1.docx"
        first = utils.temp_sibling_path(target, "part1")
        self.assertEqual(first.parent, self.dir)
        first.write_bytes(b"x")
        second = utils.temp_sibling_path(target, "part1")
        self.assertNotEqual(first, second)

    def test_safe_unlink_never_raises(self):
        utils.safe_unlink(self.dir / "tidak_ada.docx")
        utils.safe_unlink(None)

    def test_files_are_identical(self):
        a = make_docx(self.dir / "a.docx")
        b = self.dir / "b.docx"
        b.write_bytes(a.read_bytes())
        c = make_docx(self.dir / "c.docx", MINIMAL_DOCUMENT_XML.replace("Halo", "Halo dunia"))
        self.assertTrue(utils.files_are_identical(a, b))
        self.assertFalse(utils.files_are_identical(a, c))

    def test_directory_writable(self):
        ok, _ = utils.directory_is_writable(self.dir)
        self.assertTrue(ok)
        ok, reason = utils.directory_is_writable(self.dir / "tidak_ada")
        self.assertFalse(ok)
        self.assertTrue(reason)

    def test_required_space_estimate_is_conservative(self):
        self.assertGreater(validation.estimate_required_space(1_000_000), 4_000_000)


class TestModuleIntegrity(unittest.TestCase):
    """Guard against import errors and missing symbols in the non GUI modules."""

    def test_modules_import(self):
        import logger  # noqa: F401
        import split_engine  # noqa: F401
        import word_engine  # noqa: F401

    def test_word_engine_reports_absence_gracefully(self):
        import word_engine

        if os.name == "nt":
            self.skipTest("Windows menjalankan pemeriksaan COM sebenarnya.")
        available, reason = word_engine.word_is_available()
        self.assertFalse(available)
        self.assertTrue(reason)

    def test_split_engine_public_surface(self):
        import split_engine

        for symbol in (
            "analyze_document", "perform_split", "SplitRequest",
            "SplitResult", "SplitError", "OperationCancelled",
        ):
            self.assertTrue(hasattr(split_engine, symbol), symbol)


if __name__ == "__main__":
    unittest.main(verbosity=2)

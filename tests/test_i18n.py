"""Locale integrity: every language is complete, consistent and safe to render."""

from __future__ import annotations

import json
import re
import string
import sys
import unittest
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC))

import i18n  # noqa: E402
import validation  # noqa: E402

FORMATTER = string.Formatter()
SAMPLE_PARAMS = dict(
    label="label.page_a", value=5, total=4, chars=": *", name="CON", max=180, needed_bytes=1234567,
    free_bytes=1024, names="a.docx", title="Microsoft Word", start=1, end=6, page=3, count=6, a=10, b=11,
    end1=10, start2=11, n=3, actual=7, expected=6,
)


def fields(text: str) -> list:
    return sorted({name for _, name, _, _ in FORMATTER.parse(text) if name})


def load(code: str) -> dict:
    return json.loads((i18n.locales_dir() / f"{code}.json").read_text(encoding="utf-8"))


class TestLocales(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.english = load("en")

    def test_thirty_languages_with_files(self):
        self.assertEqual(len(i18n.LANGUAGES), 30)
        self.assertEqual(len({lang.code for lang in i18n.LANGUAGES}), 30)
        for lang in i18n.LANGUAGES:
            self.assertTrue((i18n.locales_dir() / f"{lang.code}.json").is_file(), lang.code)

    def test_same_keys_everywhere(self):
        for lang in i18n.LANGUAGES:
            with self.subTest(lang=lang.code):
                self.assertEqual(set(load(lang.code)), set(self.english))

    def test_placeholders_match_english(self):
        for lang in i18n.LANGUAGES:
            data = load(lang.code)
            for key, english in self.english.items():
                with self.subTest(lang=lang.code, key=key):
                    self.assertEqual(fields(data[key]), fields(english))

    def test_no_empty_or_untranslated_ui_strings(self):
        identical_allowed = {"dlg.ok", "unit.b", "unit.kb", "unit.mb", "unit.gb", "num.decimal", "label.status",
                             "dlg.details", "preview.file1", "preview.file2", "result.file1", "result.file2",
                             "header.guide", "section.pages", "section.output", "label.page_a", "label.page_b"}
        for lang in i18n.LANGUAGES:
            if lang.code == "en":
                continue
            data = load(lang.code)
            same = [k for k, v in data.items() if v == self.english[k] and k not in identical_allowed]
            with self.subTest(lang=lang.code):
                self.assertTrue(all(v.strip() for v in data.values()))
                self.assertLessEqual(len(same), 2, same)

    def test_every_string_renders(self):
        for lang in i18n.LANGUAGES:
            translator = i18n.Translator(lang.code)
            for key in self.english:
                with self.subTest(lang=lang.code, key=key):
                    text = translator.t(key, **SAMPLE_PARAMS)
                    self.assertNotIn("{", text.replace("{}", ""))

    def test_rtl_strings_avoid_constructs_tk_cannot_order(self):
        for lang in (lang for lang in i18n.LANGUAGES if lang.rtl):
            for key, value in load(lang.code).items():
                with self.subTest(lang=lang.code, key=key):
                    self.assertNotRegex(value, r"[()]")
                    self.assertNotRegex(value, r"\d\s*[-–]\s*[\d{]|\}\s*[-–]\s*[\d{]")

    def test_default_and_sample_names_are_valid_file_names(self):
        for lang in i18n.LANGUAGES:
            data = load(lang.code)
            for key in ("default.name1", "default.name2", "sample.name1", "sample.name2", "sample.file_name"):
                with self.subTest(lang=lang.code, key=key):
                    validation.validate_output_name(data[key], "label.name1")

    def test_label_parameters_are_translated(self):
        text = i18n.Translator("id").t("err.page_out_of_range", label="label.page_b", value=50, total=40)
        self.assertIn("Halaman B", text)
        self.assertIn("50", text)

    def test_size_formatting_follows_locale(self):
        self.assertEqual(i18n.Translator("fr").size(1536 * 1024), "1,50 Mo")
        self.assertEqual(i18n.Translator("en").size(1536 * 1024), "1.50 MB")

    def test_missing_parameter_never_crashes(self):
        self.assertIn("{label}", i18n.Translator("de").t("err.verify_failed"))

    def test_windows_language_tags(self):
        cases = {"id-ID": "id", "en-GB": "en", "zh-Hant-TW": "zh-TW", "zh-HK": "zh-TW", "zh-CN": "zh-CN",
                 "zh-Hans-SG": "zh-CN", "pt-PT": "pt-BR", "pt-BR": "pt-BR", "fil-PH": "fil", "jv-Latn-ID": "jv",
                 "ar-SA": "ar", "fa-IR": "fa", "ur-PK": "ur", "sr-Latn-RS": "en", None: "en", "": "en"}
        for tag, expected in cases.items():
            with self.subTest(tag=tag):
                self.assertEqual(i18n.match_tag(tag), expected)

    def test_fallback_to_english_for_unknown_code(self):
        self.assertEqual(i18n.Translator("xx").code, "en")


if __name__ == "__main__":
    unittest.main()

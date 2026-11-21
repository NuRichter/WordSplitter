"""Interface smoke tests: every language in both themes builds, fits and reacts.

These tests create real Tk windows (kept off screen) but never start Word.
"""

from __future__ import annotations

import os
import sys
import tempfile
import tkinter as tk
import unittest
from pathlib import Path
from unittest import mock

SRC = Path(__file__).resolve().parent.parent / "src"
sys.path.insert(0, str(SRC))

_TMP = tempfile.TemporaryDirectory()
os.environ["LOCALAPPDATA"] = _TMP.name  # keep the developer's settings untouched

import gui  # noqa: E402
import i18n  # noqa: E402
import theme  # noqa: E402
from settings import Settings  # noqa: E402


class GuiTest(unittest.TestCase):
    def setUp(self) -> None:
        Settings(language="en", theme="light").save()
        self.root = gui.build_root()
        self.root.geometry("+-3000+-3000")
        self.app = self.root._ws_app

    def tearDown(self) -> None:
        try:
            self.root.destroy()
        except tk.TclError:
            pass

    def stage_document(self, pages: int = 40) -> None:
        self.app.source = Path(_TMP.name) / "Thesis.docx"
        self.app.total_pages = pages
        self.app.fingerprint = [1, 1]
        self.app.file_status = "file.status.ready"
        self.app._refresh_all()

    def test_all_languages_and_themes_build(self):
        for name in ("light", "dark"):
            for language in i18n.LANGUAGES:
                with self.subTest(theme=name, lang=language.code):
                    self.app.tr = i18n.Translator(language.code)
                    self.app.mirror = gui.Mirror(language.rtl)
                    self.app.palette = theme.PALETTES[name]
                    self.app._rebuild()
                    self.root.update_idletasks()
                    self.assertEqual(self.app.btn_split.cget("text"), self.app.tr.t("btn.split"))
                    self.assertGreater(self.root.winfo_reqheight(), 200)

    def test_preview_follows_page_entries(self):
        self.stage_document(40)
        self.app.var_page_a.set("10")
        self.app.var_page_b.set("11")
        self.root.update_idletasks()
        self.assertIn("1–10", self.app.lbl_part1.cget("text"))
        self.assertIn("11–40", self.app.lbl_part2.cget("text"))
        self.app.var_page_b.set("99")
        self.assertIn("40", self.app.lbl_preview_note.cget("text"))
        self.assertEqual(self.app.lbl_part1.cget("text"), "")

    def test_controls_follow_state(self):
        self.assertIn("disabled", self.app.btn_split.state())
        self.stage_document()
        self.assertNotIn("disabled", self.app.btn_split.state())
        self.assertIn("disabled", self.app.btn_cancel.state())

    def test_language_switch_keeps_custom_names_and_translates_defaults(self):
        self.app.var_name_two.set("My chapter")
        self.app.var_language.set(i18n.BY_CODE["id"].native)
        self.app._on_language()
        self.assertEqual(self.app.var_name_one.get(), "Bagian 1")
        self.assertEqual(self.app.var_name_two.get(), "My chapter")
        self.assertEqual(Settings.load().language, "id")

    def test_theme_switch_is_saved(self):
        self.app.var_theme.set(self.app.tr.t("theme.dark"))
        self.app._on_theme()
        self.assertEqual(self.app.palette.name, "dark")
        self.assertEqual(Settings.load().theme, "dark")

    def test_window_process_never_loads_com(self):
        self.stage_document()
        self.root.update()
        for module in ("pythoncom", "pywintypes", "win32com", "word_engine"):
            self.assertNotIn(module, sys.modules)

    def test_no_timer_runs_while_idle(self):
        self.root.update()
        pending = self.root.tk.call("after", "info")
        # Only the one-shot working-set trim may be scheduled after start up.
        self.assertLessEqual(len(pending), 1)
        self.assertIsNone(self.app._poll_id)

    def test_split_validation_errors_show_dialog_not_exception(self):
        self.stage_document(10)
        self.app.var_page_a.set("3")
        self.app.var_page_b.set("3")
        with mock.patch.object(self.app, "show_error") as shown:
            self.app.on_split()
        shown.assert_called_once()
        self.assertEqual(shown.call_args[0][1]["key"], "err.page_same")

    def test_compact_mode_fits_small_screens(self):
        self.root.destroy()
        with mock.patch.object(gui.winsys, "work_area", return_value=(0, 0, 1366, 728)):
            self.root = gui.build_root(96 / 72)
            self.app = self.root._ws_app
            self.root.update_idletasks()
            self.assertTrue(self.app.compact)
            self.assertLessEqual(self.root.winfo_height(), 728)


if __name__ == "__main__":
    unittest.main()


def tearDownModule() -> None:
    _TMP.cleanup()

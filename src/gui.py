"""Tkinter user interface for WordSplitter.

The window keeps the original three step layout (source document, pages,
output) and adds only what makes that flow easier to understand: a live
picture of where the cut falls, a choice of save folder, a guide with a
practice document, thirty languages and two themes.

The interface owns no document logic. Rules live in `validation`, document
work runs in a separate worker process (`jobs`). While nothing is running the
window schedules no timers at all, so an idle WordSplitter uses no CPU.
"""

from __future__ import annotations

import logging
import os
import re
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import ttk
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

import i18n
import logger as app_logger
import theme
import utils
import validation
import winsys
from app_info import APP_NAME, VERSION
from errors import AppError, ValidationError
from settings import Settings

LOGGER = logging.getLogger("WordSplitter.gui")
POLL_MS = 80
_FILENAME_UNSAFE = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def resource_path(*parts: str) -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base.joinpath(*parts)


def safe_file_stem(text: str, fallback: str) -> str:
    cleaned = _FILENAME_UNSAFE.sub("-", text).strip(" .")
    return cleaned or fallback


class Mirror:
    """Grid helper that mirrors columns and anchors for right-to-left languages."""

    def __init__(self, rtl: bool) -> None:
        self.rtl = rtl

    def sticky(self, value: str) -> str:
        if not self.rtl:
            return value
        return value.translate(str.maketrans("ew", "we"))

    def anchor(self, value: str) -> str:
        if not self.rtl:
            return value
        return {"w": "e", "e": "w", "nw": "ne", "ne": "nw", "sw": "se", "se": "sw"}.get(value, value)

    def justify(self) -> str:
        return "right" if self.rtl else "left"

    def padx(self, value: Any) -> Any:
        if self.rtl and isinstance(value, tuple):
            return (value[1], value[0])
        return value

    def grid(self, widget: tk.Widget, row: int, column: int, columns: int, span: int = 1,
             sticky: str = "", padx: Any = 0, pady: Any = 0, **extra: Any) -> tk.Widget:
        if self.rtl:
            column = columns - column - span
        widget.grid(row=row, column=column, columnspan=span, sticky=self.sticky(sticky),
                    padx=self.padx(padx), pady=pady, **extra)
        return widget


# ---------------------------------------------------------------------------
# Dialogs
# ---------------------------------------------------------------------------


class Dialog(tk.Toplevel):
    """Themed modal dialog with localized buttons and an optional detail pane."""

    def __init__(self, app: "WordSplitterApp", title: str, message: str, *,
                 items: Sequence[str] = (), detail: Optional[str] = None,
                 buttons: Sequence[Tuple[str, str, bool]] = (), tone: str = "info",
                 extra: Optional[Callable[[ttk.Frame], None]] = None) -> None:
        super().__init__(app.root)
        self.app = app
        self.result: Optional[str] = None
        palette, s, m = app.palette, app.s, app.mirror
        self.withdraw()
        self.title(title)
        self.configure(background=palette.surface)
        self.resizable(False, False)
        self.transient(app.root)
        if app.icon_path:
            try:
                self.iconbitmap(app.icon_path)
            except tk.TclError:
                pass

        body = ttk.Frame(self, style="Dialog.TFrame", padding=(s(22), s(18), s(22), s(16)))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        wrap = s(440)

        accent = {"error": palette.danger, "warning": palette.warning, "success": palette.success}.get(tone, palette.accent)
        bar = tk.Frame(body, background=accent, height=s(3), width=s(48))
        bar.grid(row=0, column=0, sticky=m.sticky("w"), pady=(0, s(10)))
        ttk.Label(body, text=title, style="DialogTitle.TLabel", wraplength=wrap,
                  justify=m.justify(), anchor=m.anchor("w")).grid(row=1, column=0, sticky="ew")
        row = 2
        if message:
            ttk.Label(body, text=message, style="Dialog.TLabel", wraplength=wrap, justify=m.justify(),
                      anchor=m.anchor("w")).grid(row=row, column=0, sticky="ew", pady=(s(8), 0))
            row += 1
        for item in items:
            ttk.Label(body, text=item, style="Dialog.TLabel", wraplength=wrap, justify=m.justify(),
                      anchor=m.anchor("w")).grid(row=row, column=0, sticky="ew", pady=(s(4), 0))
            row += 1
        if extra is not None:
            holder = ttk.Frame(body, style="Dialog.TFrame")
            holder.grid(row=row, column=0, sticky="ew", pady=(s(10), 0))
            extra(holder)
            row += 1
        if detail:
            self._detail_text = detail
            self._detail_shown = False
            self._detail_row = row + 1
            toggle = ttk.Button(body, text=app.tr.t("dlg.details") + " ▾", command=self._toggle_detail)
            toggle.grid(row=row, column=0, sticky=m.sticky("w"), pady=(s(12), 0))
            self._detail_button = toggle
            self._detail_body = body
            row += 2

        buttons_frame = ttk.Frame(body, style="Dialog.TFrame")
        buttons_frame.grid(row=row + 1, column=0, sticky=m.sticky("e"), pady=(s(18), 0))
        ordered = list(buttons) or [(app.tr.t("dlg.ok"), "ok", True)]
        default_value = None
        widgets = []
        for text, value, primary in ordered:
            button = ttk.Button(buttons_frame, text=text, style="Accent.TButton" if primary else "TButton",
                                command=lambda v=value: self._close(v))
            widgets.append(button)
            if primary and default_value is None:
                default_value = value
                self._primary = button
        for index, button in enumerate(reversed(widgets) if m.rtl else widgets):
            button.grid(row=0, column=index, padx=(s(8) if index else 0, 0))

        cancel_value = "cancel" if any(v == "cancel" for _t, v, _p in ordered) else ordered[-1][1]
        self.protocol("WM_DELETE_WINDOW", lambda: self._close(cancel_value))
        self.bind("<Escape>", lambda _e: self._close(cancel_value))
        def on_return(event) -> None:
            if isinstance(event.widget, ttk.Button):
                event.widget.invoke()
            elif default_value is not None:
                self._close(default_value)

        self.bind("<Return>", on_return)

        self.update_idletasks()
        self._center()
        self.deiconify()
        winsys.set_dark_title_bar(self, app.palette.name == "dark")
        self.grab_set()
        try:
            (getattr(self, "_primary", None) or widgets[0]).focus_set()
        except (tk.TclError, IndexError):
            pass

    def _toggle_detail(self) -> None:
        s, m = self.app.s, self.app.mirror
        if self._detail_shown:
            self._detail_widget.destroy()
            self._detail_shown = False
            self._detail_button.configure(text=self.app.tr.t("dlg.details") + " ▾")
        else:
            text = tk.Text(self._detail_body, height=6, width=60, wrap="word", relief="flat",
                           background=self.app.palette.field, foreground=self.app.palette.muted,
                           font=("Consolas", 9), borderwidth=s(6))
            text.insert("1.0", self._detail_text)
            text.configure(state="disabled")
            text.grid(row=self._detail_row, column=0, sticky="ew", pady=(s(6), 0))
            self._detail_widget = text
            self._detail_shown = True
            self._detail_button.configure(text=self.app.tr.t("dlg.details") + " ▴")
        self.update_idletasks()

    def _center(self) -> None:
        root = self.app.root
        try:
            x = root.winfo_rootx() + (root.winfo_width() - self.winfo_reqwidth()) // 2
            y = root.winfo_rooty() + max(0, (root.winfo_height() - self.winfo_reqheight()) // 3)
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except tk.TclError:
            pass

    def _close(self, value: str) -> None:
        self.result = value
        try:
            self.grab_release()
        except tk.TclError:
            pass
        self.destroy()

    def show(self) -> Optional[str]:
        self.app.root.wait_window(self)
        return self.result


class GuideDialog(tk.Toplevel):
    """How-to, the cut rule with worked examples, and the practice document."""

    def __init__(self, app: "WordSplitterApp") -> None:
        super().__init__(app.root)
        self.app = app
        tr, s, m, palette = app.tr, app.s, app.mirror, app.palette
        self.withdraw()
        self.title(tr.t("guide.title"))
        self.configure(background=palette.surface)
        self.transient(app.root)
        if app.icon_path:
            try:
                self.iconbitmap(app.icon_path)
            except tk.TclError:
                pass

        outer = ttk.Frame(self, style="Dialog.TFrame")
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, background=palette.surface, highlightthickness=0, borderwidth=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        content = ttk.Frame(canvas, style="Dialog.TFrame", padding=(s(24), s(20)))
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="right" if m.rtl else "left", fill="both", expand=True)
        scrollbar.pack(side="left" if m.rtl else "right", fill="y")

        wrap = s(500)
        content.columnconfigure(0, weight=1)
        rows: List[tk.Widget] = []

        def label(text: str, style: str = "Dialog.TLabel", top: int = 6) -> None:
            widget = ttk.Label(content, text=text, style=style, wraplength=wrap, justify=m.justify(),
                               anchor=m.anchor("w"))
            widget.grid(row=len(rows), column=0, sticky="ew", pady=(s(top), 0))
            rows.append(widget)

        def item(marker: str, text: str, top: int = 6) -> None:
            # Marker and text in separate cells: list markers stay on the
            # leading edge in right-to-left languages, whatever Tk's bidi does.
            row = ttk.Frame(content, style="Dialog.TFrame")
            row.grid(row=len(rows), column=0, sticky="ew", pady=(s(top), 0))
            row.columnconfigure(0 if m.rtl else 1, weight=1)
            ttk.Label(row, text=marker, style="Dialog.TLabel").grid(
                row=0, column=1 if m.rtl else 0, sticky="n", padx=m.padx((0, s(8))))
            ttk.Label(row, text=text, style="Dialog.TLabel", wraplength=wrap - s(24), justify=m.justify(),
                      anchor=m.anchor("w")).grid(row=0, column=0 if m.rtl else 1, sticky="ew")
            rows.append(row)

        label(tr.t("guide.title"), "DialogTitle.TLabel", 0)
        label(tr.t("guide.intro"), top=8)
        label(tr.t("guide.steps.title"), "Strong.TLabel", 16)
        for index in range(1, 5):
            item(f"{index}.", tr.t(f"guide.step{index}"))
        label(tr.t("guide.rule.title"), "Strong.TLabel", 16)
        label(tr.t("guide.rule.body"))
        for a, b in ((10, 11), (10, 12), (10, 13)):
            plan = validation.compute_split_plan(a, b, 40)
            item("•", tr.t("guide.example", a=a, b=b, end1=plan.split_page - 1, start2=plan.split_page))

        label(tr.t("guide.sample.title"), "Strong.TLabel", 16)
        label(tr.t("guide.sample.body"))
        sample_button = ttk.Button(content, text=tr.t("guide.sample.button"), style="Accent.TButton",
                                   command=self._try_sample)
        sample_button.grid(row=len(rows), column=0, sticky=m.sticky("w"), pady=(s(10), 0))
        rows.append(sample_button)
        if app.busy:
            sample_button.state(["disabled"])

        label(tr.t("guide.tips.title"), "Strong.TLabel", 16)
        for index in range(1, 5):
            item("•", tr.t(f"guide.tip{index}"))
        label(f"{APP_NAME} {VERSION}", "DialogMuted.TLabel", 18)

        close = ttk.Button(content, text=tr.t("dlg.close"), command=self.destroy)
        close.grid(row=len(rows), column=0, sticky=m.sticky("e"), pady=(s(16), 0))

        def on_content(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def on_canvas(event) -> None:
            canvas.itemconfigure(window_id, width=event.width)

        def on_wheel(event) -> None:
            canvas.yview_scroll(int(-event.delta / 120), "units")

        content.bind("<Configure>", on_content)
        canvas.bind("<Configure>", on_canvas)
        self.bind("<MouseWheel>", on_wheel)
        self.bind("<Escape>", lambda _e: self.destroy())

        self.update_idletasks()
        width = content.winfo_reqwidth() + scrollbar.winfo_reqwidth()
        area = winsys.work_area()
        max_height = (area[3] - area[1] - s(80)) if area else s(640)
        height = min(content.winfo_reqheight(), max_height)
        canvas.configure(width=content.winfo_reqwidth(), height=height)
        x = app.root.winfo_rootx() + (app.root.winfo_width() - width) // 2
        y = (area[1] + s(20)) if area else 20
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.deiconify()
        winsys.set_dark_title_bar(self, palette.name == "dark")
        close.focus_set()

    def _try_sample(self) -> None:
        self.destroy()
        self.app.create_sample()


# ---------------------------------------------------------------------------
# Main window
# ---------------------------------------------------------------------------


class WordSplitterApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.settings = Settings.load()
        language = self.settings.language or i18n.match_tag(winsys.user_ui_language_tag())
        theme_name = self.settings.theme or ("dark" if winsys.system_prefers_dark() else "light")
        self.tr = i18n.Translator(language)
        self.palette = theme.PALETTES.get(theme_name, theme.LIGHT)
        self.scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
        self.mirror = Mirror(self.tr.rtl)
        self.icon_path: Optional[str] = None
        icon = resource_path("assets", "wordsplitter.ico")
        if icon.exists():
            self.icon_path = str(icon)
            try:
                root.iconbitmap(default=self.icon_path)
            except tk.TclError:
                self.icon_path = None

        # Document state
        self.source: Optional[Path] = None
        self.fingerprint: Optional[list] = None
        self.total_pages: Optional[int] = None
        self.file_status: Optional[str] = None
        self.output_dir: Optional[Path] = None
        self.var_page_a = tk.StringVar(master=root)
        self.var_page_b = tk.StringVar(master=root)
        self.var_name_one = tk.StringVar(master=root, value=self.tr.t("default.name1"))
        self.var_name_two = tk.StringVar(master=root, value=self.tr.t("default.name2"))
        self.var_delete = tk.BooleanVar(master=root, value=False)
        self.var_language = tk.StringVar(master=root)
        self.var_theme = tk.StringVar(master=root)
        self.status: Tuple[str, Dict[str, Any]] = ("status.select_first", {})
        self.progress_value = 0.0

        # Job state
        self.job = None
        self.job_kind: Optional[str] = None
        self._poll_id: Optional[str] = None
        self._after_analysis: Optional[Callable[[], None]] = None
        self.compact = False
        self.body: Optional[ttk.Frame] = None

        for variable in (self.var_page_a, self.var_page_b):
            variable.trace_add("write", lambda *_: self._refresh_preview())

        self._apply_look()
        self._build()
        self._fit_window(first=True)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind_all("<Control-o>", lambda e: self._main_window_key(e, self.on_select_file))
        root.bind_all("<F1>", lambda e: self._main_window_key(e, self.open_guide))
        root.bind_all("<Escape>", lambda e: self._main_window_key(e, self.on_cancel))
        # Give pages back to Windows once the window has settled.
        self._timers: set = set()
        root.bind("<Destroy>", self._on_destroy, add="+")
        self._later(1500, winsys.trim_working_set)

    # -- helpers ------------------------------------------------------------

    def _later(self, delay_ms: int, action: Callable[[], None]) -> None:
        """One-shot timer that is cancelled if the window goes away first."""
        def fire() -> None:
            self._timers.discard(timer_id)
            action()

        timer_id = self.root.after(delay_ms, fire)
        self._timers.add(timer_id)

    def _on_destroy(self, event) -> None:
        if event.widget is not self.root:
            return
        for timer_id in list(self._timers) + ([self._poll_id] if self._poll_id else []):
            try:
                self.root.after_cancel(timer_id)
            except tk.TclError:
                pass
        self._timers.clear()

    def _main_window_key(self, event, action: Callable[[], None]) -> None:
        """Shortcuts act only when the main window has focus and no dialog is up."""
        try:
            if event.widget.winfo_toplevel() is not self.root or self.root.grab_current() is not None:
                return
        except (tk.TclError, AttributeError):
            return
        action()

    def s(self, value: float) -> int:
        return int(round(value * self.scale))

    @property
    def busy(self) -> bool:
        return self.job is not None

    def _apply_look(self) -> None:
        family = theme.resolve_family(self.root, self.tr.language.font)
        self.fonts = theme.configure_fonts(self.root, family)
        theme.apply(self.root, self.palette, self.fonts, self.scale)
        winsys.set_dark_title_bar(self.root, self.palette.name == "dark")

    # -- construction -------------------------------------------------------

    def _build(self) -> None:
        tr, s, m = self.tr, self.s, self.mirror
        if self.body is not None:
            self.body.destroy()
        pad = s(10) if self.compact else s(14)
        gap = s(6) if self.compact else s(10)
        body = ttk.Frame(self.root, padding=(pad, pad, pad, s(10)))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        self.body = body

        # Header -------------------------------------------------------------
        header = ttk.Frame(body)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure((3 if m.rtl else 0), weight=1)
        titles = ttk.Frame(header)
        ttk.Label(titles, text=APP_NAME, style="Title.TLabel" if not self.compact else "AppName.TLabel",
                  anchor=m.anchor("w")).pack(anchor=m.anchor("w"))
        if not self.compact:
            ttk.Label(titles, text=tr.t("app.tagline"), style="Tagline.TLabel",
                      anchor=m.anchor("w")).pack(anchor=m.anchor("w"))
        m.grid(titles, 0, 0, 4, sticky="w")

        self.btn_guide = ttk.Button(header, text=tr.t("header.guide"), command=self.open_guide)
        m.grid(self.btn_guide, 0, 1, 4, sticky="e", padx=(s(8), 0))
        names = [language.native for language in i18n.LANGUAGES]
        self.var_language.set(tr.language.native)
        self.cmb_language = ttk.Combobox(header, textvariable=self.var_language, values=names,
                                         state="readonly", width=self._chars_for(names), justify=m.justify())
        m.grid(self.cmb_language, 0, 2, 4, sticky="e", padx=(s(8), 0))
        self.cmb_language.bind("<<ComboboxSelected>>", self._on_language)
        themes = [tr.t("theme.light"), tr.t("theme.dark")]
        self.var_theme.set(themes[1] if self.palette.name == "dark" else themes[0])
        self.cmb_theme = ttk.Combobox(header, textvariable=self.var_theme, values=themes, state="readonly",
                                      width=self._chars_for(themes), justify=m.justify())
        m.grid(self.cmb_theme, 0, 3, 4, sticky="e", padx=(s(8), 0))
        self.cmb_theme.bind("<<ComboboxSelected>>", self._on_theme)

        # 1. Source --------------------------------------------------------
        card1 = self._card(body, 1, tr.t("section.source"), gap)
        cols = 4
        card1.columnconfigure((0 if m.rtl else 3), weight=1)
        self.btn_select = ttk.Button(card1, text=tr.t("btn.select"), command=self.on_select_file)
        m.grid(self.btn_select, 1, 0, cols, sticky="w", pady=(s(6), 0))
        self.lbl_file = ttk.Label(card1, style="Strong.TLabel", anchor=m.anchor("w"), justify=m.justify())
        m.grid(self.lbl_file, 1, 1, cols, span=3, sticky="ew", padx=(s(12), 0), pady=(s(6), 0))
        self.info_values = {}
        placement = (("label.location", 2, 0, 3), ("label.size", 3, 0, 1), ("label.status", 3, 2, 1))
        for key, row, column, span in placement:
            caption = ttk.Label(card1, text=tr.t(key), style="Muted.TLabel", anchor=m.anchor("w"))
            m.grid(caption, row, column, cols, sticky="w", padx=(s(24), 0) if column else 0, pady=(s(4), 0))
            value = ttk.Label(card1, style="Card.TLabel", anchor=m.anchor("w"), justify=m.justify())
            m.grid(value, row, column + 1, cols, span=span, sticky="ew", padx=(s(12), 0), pady=(s(4), 0))
            self.info_values[key] = value

        # 2. Pages -------------------------------------------------------------
        card2 = self._card(body, 2, tr.t("section.pages"), gap)
        cols = 6
        card2.columnconfigure((1 if m.rtl else 4), weight=1)
        # In compact mode the total and "check again" share the row of Page A/B.
        total_row, pages_row = (1, 1) if self.compact else (1, 2)
        total_frame = ttk.Frame(card2, style="Dialog.TFrame")
        ttk.Label(total_frame, text=tr.t("label.total_pages"), style="Muted.TLabel").pack(
            side="right" if m.rtl else "left")
        self.lbl_total = ttk.Label(total_frame, style="Strong.TLabel")
        self.lbl_total.pack(side="right" if m.rtl else "left", padx=(s(10), s(10)))
        self.btn_reanalyze = ttk.Button(total_frame, text=tr.t("btn.reanalyze"), command=self.on_analyze)
        if self.compact:
            self.btn_reanalyze.pack(side="left" if m.rtl else "right")
            m.grid(total_frame, total_row, 4, cols, span=2, sticky="e", pady=(s(6), 0))
        else:
            spacer = ttk.Frame(total_frame, style="Dialog.TFrame")
            spacer.pack(side="right" if m.rtl else "left", fill="x", expand=True)
            self.btn_reanalyze.pack(side="left" if m.rtl else "right")
            m.grid(total_frame, total_row, 0, cols, span=6, sticky="ew", pady=(s(6), 0))

        ttk.Label(card2, text=tr.t("label.page_a"), style="Card.TLabel").grid(
            row=pages_row, column=(5 if m.rtl else 0), sticky=m.sticky("w"), pady=(s(8), 0))
        self.entry_a = ttk.Entry(card2, textvariable=self.var_page_a, width=7, justify="center")
        m.grid(self.entry_a, pages_row, 1, cols, sticky="w", padx=(s(12), 0), pady=(s(8), 0))
        ttk.Label(card2, text=tr.t("label.page_b"), style="Card.TLabel").grid(
            row=pages_row, column=(3 if m.rtl else 2), sticky=m.sticky("w"), padx=m.padx((s(20), 0)), pady=(s(8), 0))
        self.entry_b = ttk.Entry(card2, textvariable=self.var_page_b, width=7, justify="center")
        m.grid(self.entry_b, pages_row, 3, cols, sticky="w", padx=(s(12), 0), pady=(s(8), 0))
        for entry in (self.entry_a, self.entry_b):
            entry.bind("<Return>", lambda _e: self.on_split())

        self.preview = tk.Canvas(card2, height=s(10), highlightthickness=0, borderwidth=0,
                                 background=self.palette.surface)
        self.preview.grid(row=3, column=0, columnspan=cols, sticky="ew", pady=(s(10), 0))
        self.preview.bind("<Configure>", lambda _e: self._refresh_preview())
        legend = ttk.Frame(card2, style="Card.TFrame", padding=0)
        legend.configure(style="Dialog.TFrame")
        legend.grid(row=4, column=0, columnspan=cols, sticky="ew", pady=(s(4), 0))
        legend.columnconfigure(1, weight=1)
        self.lbl_part1 = ttk.Label(legend, style="Part1.TLabel")
        self.lbl_part2 = ttk.Label(legend, style="Part2.TLabel")
        self.lbl_preview_note = ttk.Label(legend, style="Muted.TLabel", wraplength=s(560), justify=m.justify())
        self.lbl_part1.grid(row=0, column=2 if m.rtl else 0, sticky=m.sticky("w"))
        self.lbl_part2.grid(row=0, column=0 if m.rtl else 2, sticky=m.sticky("e"))
        self.lbl_preview_note.grid(row=1, column=0, columnspan=3, sticky=m.sticky("w"))
        if not self.compact:
            self.lbl_hint = ttk.Label(card2, text=tr.t("hint.pages"), style="Muted.TLabel",
                                      justify=m.justify(), anchor=m.anchor("w"))
            self.lbl_hint.grid(row=5, column=0, columnspan=cols, sticky="ew", pady=(s(6), 0))
            self._wrap_with(card2, self.lbl_hint)

        # 3. Output ------------------------------------------------------------
        card3 = self._card(body, 3, tr.t("section.output"), gap)
        cols = 3
        card3.columnconfigure(1, weight=1)
        self.entries_names = []
        for row, (key, variable) in enumerate((("label.name1", self.var_name_one),
                                                ("label.name2", self.var_name_two)), start=1):
            ttk.Label(card3, text=tr.t(key), style="Card.TLabel").grid(
                row=row, column=(2 if m.rtl else 0), sticky=m.sticky("w"), pady=(s(6 if row == 1 else 4), 0))
            entry = ttk.Entry(card3, textvariable=variable, justify=m.justify())
            entry.grid(row=row, column=1, sticky="ew", padx=(s(12), s(4)), pady=(s(6 if row == 1 else 4), 0))
            entry.bind("<Return>", lambda _e: self.on_split())
            self.entries_names.append(entry)
            ttk.Label(card3, text=".docx", style="Muted.TLabel").grid(
                row=row, column=(0 if m.rtl else 2), sticky=m.sticky("w"), pady=(s(6 if row == 1 else 4), 0))
        ttk.Label(card3, text=tr.t("label.save_to"), style="Card.TLabel").grid(
            row=3, column=(2 if m.rtl else 0), sticky=m.sticky("w"), pady=(s(8), 0))
        self.lbl_folder = ttk.Label(card3, style="Card.TLabel", anchor=m.anchor("w"), justify=m.justify())
        self.lbl_folder.grid(row=3, column=1, sticky="ew", padx=(s(12), s(4)), pady=(s(8), 0))
        self.btn_folder = ttk.Button(card3, text=tr.t("btn.change"), command=self.on_choose_folder)
        self.btn_folder.grid(row=3, column=(0 if m.rtl else 2), sticky=m.sticky("e"), pady=(s(8), 0))
        self.chk_delete = ttk.Checkbutton(card3, text=tr.t("label.delete_original"), variable=self.var_delete)
        self.chk_delete.grid(row=4, column=0, columnspan=cols, sticky=m.sticky("w"), pady=(s(8), 0))

        # Actions ----------------------------------------------------------
        actions = ttk.Frame(body)
        actions.grid(row=4, column=0, sticky="ew", pady=(gap + s(2), 0))
        cols = 3
        actions.columnconfigure((2 if m.rtl else 0), weight=1)
        self.progress = ttk.Progressbar(actions, mode="determinate", maximum=100, value=self.progress_value)
        m.grid(self.progress, 0, 0, cols, sticky="ew", padx=(0, s(12)))
        self.btn_cancel = ttk.Button(actions, text=tr.t("btn.cancel"), command=self.on_cancel)
        m.grid(self.btn_cancel, 0, 1, cols, padx=(0, s(8)))
        self.btn_split = ttk.Button(actions, text=tr.t("btn.split"), style="Accent.TButton", command=self.on_split)
        m.grid(self.btn_split, 0, 2, cols)

        self.lbl_status = ttk.Label(body, style="Status.TLabel", justify=m.justify(), anchor=m.anchor("w"))
        self.lbl_status.grid(row=5, column=0, sticky="ew", pady=(s(8), 0))
        self._wrap_with(body, self.lbl_status, s(2 * 16))
        if not self.compact:
            log_dir = app_logger.log_file_path()
            self.lbl_log = ttk.Label(body, style="Link.TLabel", cursor="hand2", anchor=m.anchor("w"),
                                     text=f"{tr.t('status.log')}: {log_dir.parent if log_dir else '-'}")
            self.lbl_log.grid(row=6, column=0, sticky="ew", pady=(s(4), 0))
            self.lbl_log.bind("<Button-1>", lambda _e: self._open_logs())

        self._refresh_all()

    def _chars_for(self, texts: Sequence[str]) -> int:
        """Combobox width (in average characters) that fits the widest entry."""
        from tkinter import font as tkfont

        measure = tkfont.nametofont("TkTextFont", root=self.root)
        unit = max(1, measure.measure("0"))
        widest = max((measure.measure(text) for text in texts), default=unit * 6)
        return max(6, int(widest / unit) + 2)

    def _card(self, parent: ttk.Frame, number: int, title: str, gap: int) -> ttk.Frame:
        s, m = self.s, self.mirror
        card = ttk.Frame(parent, style="Card.TFrame", padding=(s(14), s(8), s(14), s(10)))
        card.grid(row=number, column=0, sticky="ew", pady=(gap, 0))
        heading = ttk.Label(card, text=title, style="Section.TLabel", anchor=m.anchor("w"))
        heading.grid(row=0, column=0, columnspan=8, sticky=m.sticky("w"))
        return card

    def _wrap_with(self, container: tk.Widget, label: ttk.Label, margin: int = 0) -> None:
        def update(event) -> None:
            label.configure(wraplength=max(self.s(200), event.width - margin - self.s(8)))
        container.bind("<Configure>", update, add="+")

    # -- refresh ------------------------------------------------------------

    def _refresh_all(self) -> None:
        tr = self.tr
        if self.source is None:
            self.lbl_file.configure(text=tr.t("file.none"))
            for value in self.info_values.values():
                value.configure(text="–")
        else:
            self.lbl_file.configure(text=self.source.name)
            self.info_values["label.location"].configure(text=str(self.source.parent))
            try:
                size = tr.size(self.source.stat().st_size)
            except OSError:
                size = "–"
            self.info_values["label.size"].configure(text=size)
            self.info_values["label.status"].configure(text=tr.t(self.file_status) if self.file_status else "–")
        self.lbl_total.configure(text=str(self.total_pages) if self.total_pages else "–")
        folder = self.output_dir or (self.source.parent if self.source else None)
        self.lbl_folder.configure(text=str(folder) if folder else "–")
        self._set_status(*self.status)
        self._refresh_preview()
        self._update_controls()
        self._wrap_paths()

    def _wrap_paths(self) -> None:
        width = self.s(430)
        for label in (self.info_values.get("label.location"), self.lbl_folder, self.lbl_file):
            if label is not None:
                label.configure(wraplength=width)

    def _set_status(self, key: str, params: Optional[Dict[str, Any]] = None) -> None:
        self.status = (key, params or {})
        if self.body is not None:
            self.lbl_status.configure(text=self.tr.t(key, **(params or {})))

    def _update_controls(self) -> None:
        busy = self.busy
        has_file = self.source is not None
        ready = has_file and self.total_pages is not None

        def toggle(widget: ttk.Widget, enabled: bool) -> None:
            widget.state(["!disabled"] if enabled else ["disabled"])

        toggle(self.btn_select, not busy)
        toggle(self.btn_reanalyze, has_file and not busy)
        toggle(self.btn_split, ready and not busy)
        toggle(self.btn_cancel, busy and not (self.job is not None and self.job.committing))
        toggle(self.btn_folder, has_file and not busy)
        toggle(self.chk_delete, not busy)
        for entry in (self.entry_a, self.entry_b, *self.entries_names):
            toggle(entry, not busy)
        self.cmb_language.configure(state="disabled" if busy else "readonly")
        self.cmb_theme.configure(state="disabled" if busy else "readonly")

    def _refresh_preview(self) -> None:
        if self.body is None:
            return
        tr, p, m = self.tr, self.palette, self.mirror
        canvas = self.preview
        width = max(1, canvas.winfo_width())
        height = max(1, canvas.winfo_height())
        canvas.delete("all")
        self.lbl_part1.configure(text="")
        self.lbl_part2.configure(text="")
        note, danger = "", False
        plan = None
        if self.total_pages is None:
            note = tr.t("preview.need_file")
        else:
            raw_a, raw_b = self.var_page_a.get().strip(), self.var_page_b.get().strip()
            if not raw_a or not raw_b:
                note = tr.t("preview.waiting")
            else:
                try:
                    a = validation.parse_page_number(raw_a, "label.page_a")
                    b = validation.parse_page_number(raw_b, "label.page_b")
                    plan = validation.compute_split_plan(a, b, self.total_pages)
                except ValidationError as exc:
                    note, danger = tr.message(exc.to_dict()), True
        if plan is None:
            canvas.create_rectangle(0, 0, width, height, fill=p.border, width=0)
            self.lbl_preview_note.configure(text=note, style="Danger.TLabel" if danger else "Muted.TLabel")
            self.lbl_preview_note.grid()
            return
        self.lbl_preview_note.configure(text="")
        self.lbl_preview_note.grid_remove()
        fraction = (plan.split_page - 1) / plan.total_pages
        cut = max(2, min(width - 2, int(round(width * fraction))))
        first, second = (p.part1, p.part2)
        if m.rtl:
            canvas.create_rectangle(width - cut, 0, width, height, fill=first, width=0)
            canvas.create_rectangle(0, 0, width - cut - 2, height, fill=second, width=0)
        else:
            canvas.create_rectangle(0, 0, cut, height, fill=first, width=0)
            canvas.create_rectangle(cut + 2, 0, width, height, fill=second, width=0)

        def describe(start: int, end: int) -> str:
            return tr.t("preview.single", page=start) if start == end else tr.t("preview.range", start=start, end=end)

        s1, e1 = plan.part_one
        s2, e2 = plan.part_two
        self.lbl_part1.configure(text=f"{tr.t('preview.file1')} · {describe(s1, e1)}")
        self.lbl_part2.configure(text=f"{tr.t('preview.file2')} · {describe(s2, e2)}")

    # -- window sizing --------------------------------------------------------

    def _fit_window(self, first: bool = False) -> None:
        root = self.root
        root.update_idletasks()
        req_w, req_h = root.winfo_reqwidth(), root.winfo_reqheight()
        area = winsys.work_area() or (0, 0, root.winfo_screenwidth(), root.winfo_screenheight())
        avail_w = area[2] - area[0] - self.s(16)
        avail_h = area[3] - area[1] - self.s(48)  # title bar and a little air
        if req_h > avail_h and not self.compact:
            self.compact = True
            self._build()
            root.update_idletasks()
            req_w, req_h = root.winfo_reqwidth(), root.winfo_reqheight()
        width = min(max(req_w, self.s(640)), avail_w)
        height = min(req_h, avail_h)
        root.minsize(min(req_w, avail_w), min(req_h, avail_h))
        if first:
            x = area[0] + (area[2] - area[0] - width) // 2
            y = area[1] + max(0, (area[3] - area[1] - height - self.s(32)) // 2)
            root.geometry(f"{width}x{height}+{x}+{y}")
        else:
            root.geometry(f"{max(width, root.winfo_width())}x{max(height, min(root.winfo_height(), avail_h))}")

    # -- language and theme ---------------------------------------------------

    def _on_language(self, _event=None) -> None:
        native = self.var_language.get()
        language = next((lang for lang in i18n.LANGUAGES if lang.native == native), None)
        if language is None or language.code == self.tr.code or self.busy:
            return
        old = self.tr
        self.tr = i18n.Translator(language.code)
        self.mirror = Mirror(self.tr.rtl)
        for variable, key in ((self.var_name_one, "default.name1"), (self.var_name_two, "default.name2")):
            if variable.get().strip() == old.t(key):
                variable.set(self.tr.t(key))
        self.settings.language = language.code
        self.settings.save()
        LOGGER.info("Language set to %s", language.code)
        self._rebuild()

    def _on_theme(self, _event=None) -> None:
        name = "dark" if self.var_theme.get() == self.tr.t("theme.dark") else "light"
        if name == self.palette.name or self.busy:
            return
        self.palette = theme.PALETTES[name]
        self.settings.theme = name
        self.settings.save()
        self._rebuild()

    def _rebuild(self) -> None:
        self.compact = False
        self._apply_look()
        self._build()
        self._fit_window()
        self.root.update_idletasks()
        winsys.set_dark_title_bar(self.root, self.palette.name == "dark")

    # -- dialogs --------------------------------------------------------------

    def dialog(self, title_key: str, message: str, **kwargs: Any) -> Optional[str]:
        return Dialog(self, self.tr.t(title_key), message, **kwargs).show()

    def show_error(self, title_key: str, error: Dict[str, Any]) -> None:
        tr = self.tr
        detail = error.get("detail")
        items = []
        if error.get("key") not in ("err.page_empty", "err.page_not_int", "err.page_zero", "err.page_same",
                                    "err.page_out_of_range", "err.names_same") and detail:
            items.append(tr.t("dlg.log_hint"))
        self.dialog(title_key, tr.message(error), items=items, detail=detail, tone="error",
                    buttons=[(tr.t("dlg.ok"), "ok", True)])

    def open_guide(self) -> None:
        GuideDialog(self)

    def _open_logs(self) -> None:
        path = app_logger.log_file_path()
        if path is not None:
            try:
                os.startfile(str(path.parent))  # type: ignore[attr-defined]
            except OSError:
                pass

    # -- actions ----------------------------------------------------------------

    def on_select_file(self) -> None:
        if self.busy:
            return
        from tkinter import filedialog

        initial = self.settings.last_folder if self.settings.last_folder and os.path.isdir(self.settings.last_folder) else None
        selected = filedialog.askopenfilename(
            parent=self.root, title=self.tr.t("dlg.pick_file"), initialdir=initial,
            filetypes=[(self.tr.t("dlg.filetype"), "*.docx")],
        )
        if selected:
            self.load_source(selected)

    def load_source(self, raw_path: str, keep_inputs: bool = False) -> None:
        try:
            path = validation.validate_input_file(raw_path)
        except ValidationError as exc:
            LOGGER.info("Rejected file %s: %s", raw_path, exc)
            self.show_error("dlg.file_error.title", exc.to_dict())
            return
        self.source = path
        self.fingerprint = None
        self.total_pages = None
        self.output_dir = None
        if not keep_inputs:
            self.var_page_a.set("")
            self.var_page_b.set("")
        self.settings.last_folder = str(path.parent)
        self.settings.save()
        LOGGER.info("Selected %s", path)
        self.on_analyze()

    def on_analyze(self, status_key: str = "status.analyzing") -> None:
        if self.busy or self.source is None:
            return
        self.total_pages = None
        self.fingerprint = None
        self.file_status = "file.status.analyzing"
        self._start_job({"kind": "analyze", "source": str(self.source)}, status_key)

    def on_choose_folder(self) -> None:
        if self.busy or self.source is None:
            return
        from tkinter import filedialog

        current = self.output_dir or self.source.parent
        selected = filedialog.askdirectory(parent=self.root, title=self.tr.t("dlg.pick_folder"),
                                           initialdir=str(current), mustexist=True)
        if selected:
            self.output_dir = Path(selected)
            self._refresh_all()

    def on_split(self) -> None:
        if self.busy or self.source is None or self.total_pages is None:
            return
        tr = self.tr
        directory = self.output_dir or self.source.parent
        try:
            page_a = validation.parse_page_number(self.var_page_a.get(), "label.page_a")
            page_b = validation.parse_page_number(self.var_page_b.get(), "label.page_b")
            validation.compute_split_plan(page_a, page_b, self.total_pages)
            output_one, output_two = validation.resolve_output_paths(
                self.source, directory, self.var_name_one.get(), self.var_name_two.get())
        except ValidationError as exc:
            self.show_error("dlg.input_error.title", exc.to_dict())
            return

        overwrite = False
        conflicts = validation.existing_outputs(output_one, output_two)
        if conflicts:
            choice = self.dialog("dlg.conflict.title", tr.t("dlg.conflict.body"),
                                 items=[f"•  {p.name}" for p in conflicts], tone="warning",
                                 buttons=[(tr.t("dlg.conflict.replace"), "replace", False),
                                          (tr.t("dlg.conflict.rename"), "rename", True),
                                          (tr.t("dlg.cancel"), "cancel", False)])
            if choice == "rename":
                self._set_status("status.rename")
                (self.entries_names[0] if output_one in conflicts else self.entries_names[1]).focus_set()
                return
            if choice != "replace":
                return
            overwrite = True

        if self.var_delete.get():
            choice = self.dialog("dlg.delete.title", tr.t("dlg.delete.body"), items=[f"•  {self.source.name}"],
                                 tone="warning", buttons=[(tr.t("dlg.delete.confirm"), "yes", True),
                                                          (tr.t("dlg.cancel"), "cancel", False)])
            if choice != "yes":
                return

        request = {
            "kind": "split",
            "source": str(self.source),
            "fingerprint": self.fingerprint,
            "total_pages": self.total_pages,
            "page_a": page_a,
            "page_b": page_b,
            "output_one": str(output_one),
            "output_two": str(output_two),
            "delete_original": bool(self.var_delete.get()),
            "overwrite_allowed": overwrite,
        }
        LOGGER.info("Split requested: %s", {k: v for k, v in request.items() if k != "fingerprint"})
        self._start_job(request, "status.splitting")

    def on_cancel(self) -> None:
        if self.job is None:
            return
        if self.job.cancel():
            self._set_status("status.cancelling")
            self.btn_cancel.state(["disabled"])

    def create_sample(self) -> None:
        if self.busy:
            return
        tr = self.tr
        stem = safe_file_stem(tr.t("sample.file_name"), "WordSplitter sample")
        candidates = []
        documents = winsys.documents_dir()
        if documents:
            candidates.append(Path(documents) / "WordSplitter")
        candidates.append(app_logger.app_data_dir() / "samples")
        import sample_doc

        created = None
        error = None
        for folder in candidates:
            for index in range(0, 20):
                target = folder / (f"{stem}.docx" if index == 0 else f"{stem} ({index + 1}).docx")
                try:
                    created = sample_doc.build(target, tr.t, font=tr.language.font if tr.language.font != "Segoe UI" else "Calibri",
                                               rtl=tr.rtl, lang=tr.code)
                    break
                except PermissionError as exc:  # e.g. the previous sample is open in Word
                    error = exc
                    continue
                except OSError as exc:
                    error = exc
                    break
            if created:
                break
        if created is None:
            self.show_error("dlg.error.title", {"key": "err.sample_failed", "detail": str(error)})
            return
        LOGGER.info("Sample created at %s", created)
        self.var_name_one.set(safe_file_stem(tr.t("sample.name1"), "Sample part 1"))
        self.var_name_two.set(safe_file_stem(tr.t("sample.name2"), "Sample part 2"))
        self.var_delete.set(False)
        self.var_page_a.set(str(sample_doc.SPLIT_A))
        self.var_page_b.set(str(sample_doc.SPLIT_B))
        self._after_analysis = lambda: self._set_status("status.sample_ready")
        self.load_source(str(created), keep_inputs=True)

    # -- jobs -------------------------------------------------------------------

    def _start_job(self, request: Dict[str, Any], status_key: str) -> None:
        import jobs

        try:
            self.job = jobs.WorkerJob(request)
        except Exception as exc:  # noqa: BLE001 - spawning can fail on locked down systems
            LOGGER.exception("Could not start the worker")
            self.job = None
            self.show_error("dlg.error.title", {"key": "err.worker_crashed", "detail": f"{type(exc).__name__}: {exc}"})
            return
        self.job_kind = request["kind"]
        self.progress_value = 0
        self.progress.configure(value=0)
        self._set_status(status_key)
        self._refresh_all()
        self._poll_id = self.root.after(POLL_MS, self._poll)

    def _poll(self) -> None:
        self._poll_id = None
        job = self.job
        if job is None:
            return
        import queue

        try:
            while True:
                kind, data = job.events.get_nowait()
                self._handle(kind, data)
                if self.job is None:
                    return
        except queue.Empty:
            pass
        self._poll_id = self.root.after(POLL_MS, self._poll)

    def _finish_job(self) -> str:
        kind = self.job_kind or ""
        self.job = None
        self.job_kind = None
        if self._poll_id is not None:
            self.root.after_cancel(self._poll_id)
            self._poll_id = None
        self.progress_value = 0
        self.progress.configure(value=0)
        self._update_controls()
        # The worker did the heavy lifting; let Windows reclaim what the GUI touched.
        self._later(800, winsys.trim_working_set)
        return kind

    def _handle(self, kind: str, data: Dict[str, Any]) -> None:
        if kind == "progress":
            self.progress_value = float(data.get("pct") or 0)
            self.progress.configure(value=self.progress_value)
            self._set_status(str(data.get("key") or "status.splitting"))
            return
        if kind == "commit":
            self.btn_cancel.state(["disabled"])
            return
        if kind not in ("done", "error", "cancelled"):
            return
        job_kind = self._finish_job()
        if kind != "done" or job_kind != "analyze":
            self._after_analysis = None
        if kind == "cancelled":
            if job_kind == "analyze":
                self.file_status = None
            self._set_status("status.cancelled")
            self._refresh_all()
            return
        if kind == "error":
            error = data.get("error") or {"key": "err.unexpected"}
            LOGGER.warning("%s failed: %s", job_kind, error)
            if error.get("key") == "err.source_changed" and self.source is not None:
                self.file_status = "file.status.changed"
                self.on_analyze("status.reanalyzing")
                return
            if job_kind == "analyze":
                self.file_status = "file.status.failed"
                self._set_status("status.failed")
                self._refresh_all()
                self.show_error("dlg.analysis_error.title", error)
            else:
                self._set_status("status.failed")
                self._refresh_all()
                self.show_error("dlg.error.title", error)
            return
        result = data.get("result") or {}
        if job_kind == "analyze":
            self.total_pages = int(result.get("pages") or 0) or None
            self.fingerprint = result.get("fingerprint")
            self.file_status = "file.status.ready"
            if self.total_pages and self.total_pages >= 2 and not self.var_page_a.get().strip() and not self.var_page_b.get().strip():
                middle = max(1, self.total_pages // 2)
                self.var_page_a.set(str(middle))
                self.var_page_b.set(str(middle + 1))
            self._set_status("status.analyzed")
            self._refresh_all()
            if self._after_analysis is not None:
                callback, self._after_analysis = self._after_analysis, None
                callback()
            return
        self._show_success(result)

    def _show_success(self, result: Dict[str, Any]) -> None:
        tr = self.tr
        warnings = result.get("warnings") or []
        output_one = Path(result.get("output_one", ""))
        output_two = Path(result.get("output_two", ""))
        plan = result.get("plan") or [0, 0, 1, 1]
        split_page, total = int(plan[2]), int(plan[3])

        def describe(start: int, end: int) -> str:
            return tr.t("preview.single", page=start) if start == end else tr.t("preview.range", start=start, end=end)

        items = [
            f"{tr.t('result.file1')}: {output_one.name}",
            f"    {describe(1, split_page - 1)}  ·  {tr.t('result.pages', count=result.get('pages_one'))}",
            f"{tr.t('result.file2')}: {output_two.name}",
            f"    {describe(split_page, total)}  ·  {tr.t('result.pages', count=result.get('pages_two'))}",
        ]
        removed = result.get("original_removed", "kept")
        items.append({"recycled": tr.t("result.recycled"), "deleted": tr.t("result.deleted")}.get(removed, tr.t("result.kept")))
        for warning in warnings:
            items.append("⚠  " + tr.message(warning))

        if removed in ("recycled", "deleted"):
            self.source = None
            self.total_pages = None
            self.fingerprint = None
            self.file_status = None
        self._set_status("status.done_notes" if warnings else "status.done")
        self._refresh_all()

        choice = self.dialog("dlg.notes.title" if warnings else "dlg.success.title", tr.t("dlg.success.body"),
                             items=items, tone="warning" if warnings else "success",
                             buttons=[(tr.t("dlg.open_folder"), "open", False), (tr.t("dlg.close"), "close", True)])
        if choice == "open":
            try:
                subprocess.Popen(["explorer", "/select,", str(output_one)])
            except OSError:
                pass

    # -- shutdown ---------------------------------------------------------------

    def on_close(self) -> None:
        if self.busy:
            choice = self.dialog("dlg.busy.title", self.tr.t("dlg.busy.body"), tone="warning",
                                 buttons=[(self.tr.t("dlg.busy.confirm"), "stop", True),
                                          (self.tr.t("dlg.cancel"), "cancel", False)])
            if choice != "stop":
                return
            if self.job is not None:
                self.job.cancel()
        LOGGER.info("Window closed by the user.")
        self.root.destroy()


def build_root(scaling: Optional[float] = None) -> tk.Tk:
    root = tk.Tk(className=APP_NAME)
    root.withdraw()
    if scaling:
        root.tk.call("tk", "scaling", scaling)  # used by the screenshot tool to emulate other DPIs
    root.title(APP_NAME)
    app = WordSplitterApp(root)
    root._ws_app = app  # type: ignore[attr-defined]
    root.deiconify()
    root.update_idletasks()
    winsys.set_dark_title_bar(root, app.palette.name == "dark")
    return root


def run() -> None:
    root = build_root()
    root.mainloop()

"""Two colour themes on top of ttk's "clam" engine.

"clam" is drawn by Tk itself, so unlike the native "vista" theme every colour
can be set, and both themes look the same on Windows 10 and 11. Colours follow
the Windows 11 neutrals with Word's blue as the accent.
"""

from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass
from tkinter import font as tkfont
from tkinter import ttk


@dataclass(frozen=True)
class Palette:
    name: str
    window: str
    surface: str
    border: str
    text: str
    muted: str
    accent: str
    accent_hover: str
    accent_text: str
    button: str
    button_hover: str
    field: str
    field_border: str
    disabled: str
    select: str
    part1: str
    part2: str
    danger: str
    warning: str
    success: str


LIGHT = Palette(
    name="light", window="#f3f3f3", surface="#ffffff", border="#d9d9d9", text="#1b1b1b",
    muted="#5c5c5c", accent="#185abd", accent_hover="#1f6ad6", accent_text="#ffffff",
    button="#fbfbfb", button_hover="#eef3fb", field="#ffffff", field_border="#bdbdbd",
    disabled="#a0a0a0", select="#cfe0f8", part1="#185abd", part2="#0e7c66",
    danger="#c42b1c", warning="#9d5d00", success="#0f7b0f",
)
DARK = Palette(
    name="dark", window="#202020", surface="#2b2b2b", border="#3d3d3d", text="#f3f3f3",
    muted="#b4b4b4", accent="#4c94f0", accent_hover="#68a6f5", accent_text="#0b1a33",
    button="#373737", button_hover="#414b5a", field="#1f1f1f", field_border="#5a5a5a",
    disabled="#7a7a7a", select="#27466f", part1="#4c94f0", part2="#2fbf9f",
    danger="#ff99a4", warning="#fce100", success="#6ccb5f",
)
PALETTES = {"light": LIGHT, "dark": DARK}


def resolve_family(root: tk.Misc, wanted: str) -> str:
    """Return `wanted` if the font is installed, otherwise Segoe UI."""
    try:
        actual = tkfont.Font(root=root, family=wanted, size=10).actual("family")
        if actual.lower() == wanted.lower():
            return wanted
    except tk.TclError:
        pass
    return "Segoe UI"


def configure_fonts(root: tk.Misc, family: str) -> dict:
    """Point Tk's named fonts at the language's UI font. Returns extra fonts."""
    for name, size, weight in (
        ("TkDefaultFont", 10, "normal"),
        ("TkTextFont", 10, "normal"),
        ("TkMenuFont", 10, "normal"),
        ("TkHeadingFont", 10, "bold"),
        ("TkCaptionFont", 10, "normal"),
        ("TkTooltipFont", 9, "normal"),
        ("TkFixedFont", 10, "normal"),
    ):
        try:
            tkfont.nametofont(name, root=root).configure(family=family, size=size, weight=weight)
        except tk.TclError:
            pass
    return {
        "title": (family, 15, "bold"),
        "section": (family, 10, "bold"),
        "strong": (family, 10, "bold"),
        "small": (family, 9),
        "big": (family, 12, "bold"),
    }


def apply(root: tk.Tk, palette: Palette, fonts: dict, scale: float) -> None:
    """Configure every ttk style the interface uses."""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        return
    p = palette
    root.configure(background=p.window)

    style.configure(".", background=p.window, foreground=p.text, fieldbackground=p.field,
                    bordercolor=p.border, lightcolor=p.window, darkcolor=p.window,
                    troughcolor=p.border, focuscolor=p.accent, selectbackground=p.select,
                    selectforeground=p.text, insertcolor=p.text, font="TkDefaultFont")
    style.map(".", foreground=[("disabled", p.disabled)])

    style.configure("TFrame", background=p.window)
    style.configure("Card.TFrame", background=p.surface, bordercolor=p.border, relief="solid", borderwidth=1)
    style.configure("TLabel", background=p.window, foreground=p.text)
    style.configure("Card.TLabel", background=p.surface, foreground=p.text)
    style.configure("Muted.TLabel", background=p.surface, foreground=p.muted)
    style.configure("Small.TLabel", background=p.window, foreground=p.muted, font=fonts["small"])
    style.configure("Title.TLabel", background=p.window, foreground=p.text, font=fonts["title"])
    style.configure("AppName.TLabel", background=p.window, foreground=p.text, font=fonts["big"])
    style.configure("Tagline.TLabel", background=p.window, foreground=p.muted)
    style.configure("Section.TLabel", background=p.surface, foreground=p.text, font=fonts["section"])
    style.configure("Strong.TLabel", background=p.surface, foreground=p.text, font=fonts["strong"])
    style.configure("Status.TLabel", background=p.window, foreground=p.text)
    style.configure("Danger.TLabel", background=p.surface, foreground=p.danger)
    style.configure("Part1.TLabel", background=p.surface, foreground=p.part1, font=fonts["strong"])
    style.configure("Part2.TLabel", background=p.surface, foreground=p.part2, font=fonts["strong"])
    style.configure("Dialog.TFrame", background=p.surface)
    style.configure("Dialog.TLabel", background=p.surface, foreground=p.text)
    style.configure("DialogTitle.TLabel", background=p.surface, foreground=p.text, font=fonts["big"])
    style.configure("DialogMuted.TLabel", background=p.surface, foreground=p.muted)

    style.configure("TButton", background=p.button, foreground=p.text, bordercolor=p.field_border,
                    lightcolor=p.button, darkcolor=p.button, focuscolor=p.accent,
                    padding=(round(12 * scale), round(4 * scale)), relief="flat")
    style.map("TButton",
              background=[("disabled", p.surface), ("pressed", p.select), ("active", p.button_hover)],
              foreground=[("disabled", p.disabled)],
              bordercolor=[("focus", p.accent)])
    style.configure("Accent.TButton", background=p.accent, foreground=p.accent_text, bordercolor=p.accent,
                    lightcolor=p.accent, darkcolor=p.accent, font=fonts["strong"],
                    padding=(round(20 * scale), round(5 * scale)))
    style.map("Accent.TButton",
              background=[("disabled", p.border), ("pressed", p.accent), ("active", p.accent_hover)],
              foreground=[("disabled", p.disabled)],
              bordercolor=[("disabled", p.border), ("focus", p.text)])

    for widget in ("TEntry", "TCombobox"):
        style.configure(widget, fieldbackground=p.field, foreground=p.text, bordercolor=p.field_border,
                        lightcolor=p.field, darkcolor=p.field, insertcolor=p.text, padding=round(3 * scale),
                        arrowcolor=p.text, background=p.button)
        style.map(widget,
                  bordercolor=[("focus", p.accent)],
                  lightcolor=[("focus", p.accent)],
                  fieldbackground=[("readonly", p.field), ("disabled", p.surface)],
                  foreground=[("disabled", p.disabled)],
                  selectbackground=[("readonly", p.field)],
                  selectforeground=[("readonly", p.text)])
    root.option_add("*TCombobox*Listbox.background", p.field)
    root.option_add("*TCombobox*Listbox.foreground", p.text)
    root.option_add("*TCombobox*Listbox.selectBackground", p.select)
    root.option_add("*TCombobox*Listbox.selectForeground", p.text)
    root.option_add("*TCombobox*Listbox.font", "TkDefaultFont")

    style.configure("TCheckbutton", background=p.surface, foreground=p.text, indicatorbackground=p.field,
                    indicatorforeground=p.accent_text, bordercolor=p.field_border, focuscolor=p.surface,
                    indicatormargin=(0, 0, round(6 * scale), 0))
    style.map("TCheckbutton",
              indicatorbackground=[("selected", p.accent), ("pressed", p.select)],
              background=[("active", p.surface)],
              foreground=[("disabled", p.disabled)])

    style.configure("Horizontal.TProgressbar", troughcolor=p.border, background=p.accent,
                    bordercolor=p.border, lightcolor=p.accent, darkcolor=p.accent,
                    thickness=max(4, round(6 * scale)))
    style.configure("TSeparator", background=p.border)
    style.configure("Vertical.TScrollbar", background=p.button, troughcolor=p.surface, bordercolor=p.surface,
                    arrowcolor=p.text, lightcolor=p.button, darkcolor=p.button)
    style.map("Vertical.TScrollbar", background=[("active", p.button_hover)])
    style.configure("Link.TLabel", background=p.window, foreground=p.accent, font=fonts["small"])

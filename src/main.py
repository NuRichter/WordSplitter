"""WordSplitter entry point.

Responsibilities: configure logging, install a global exception hook so that no
traceback ever reaches the user, enable per monitor DPI awareness on Windows,
and start the GUI.
"""

from __future__ import annotations

import ctypes
import os
import sys
import threading
import traceback
from pathlib import Path

# Allow both `python src/main.py` and the PyInstaller one file bundle to resolve
# the sibling modules without requiring a package installation.
sys.path.insert(0, str(Path(__file__).resolve().parent))

import logger as app_logger  # noqa: E402


def _enable_dpi_awareness() -> None:
    if os.name != "nt":
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # noqa: BLE001 - purely cosmetic
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:  # noqa: BLE001
            pass


def _install_exception_hooks(log) -> None:
    def hook(exc_type, exc_value, exc_tb) -> None:
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        log.critical(
            "Unhandled exception:\n%s",
            "".join(traceback.format_exception(exc_type, exc_value, exc_tb)),
        )
        _show_fatal_dialog(exc_type.__name__)

    sys.excepthook = hook

    def thread_hook(args) -> None:
        log.critical(
            "Unhandled exception in thread %s:\n%s",
            args.thread.name if args.thread else "unknown",
            "".join(
                traceback.format_exception(args.exc_type, args.exc_value, args.exc_traceback)
            ),
        )

    threading.excepthook = thread_hook


def _show_fatal_dialog(exception_name: str) -> None:
    """Report a fatal error without ever exposing a raw traceback."""
    log_path = app_logger.log_file_path()
    message = (
        "WordSplitter mengalami kesalahan yang tidak dapat dipulihkan dan harus ditutup.\n\n"
        f"Jenis kesalahan: {exception_name}\n"
    )
    if log_path:
        message += f"\nDetail teknis telah disimpan pada:\n{log_path}"
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("WordSplitter", message)
        root.destroy()
    except Exception:  # noqa: BLE001 - last resort
        try:
            if os.name == "nt":
                ctypes.windll.user32.MessageBoxW(0, message, "WordSplitter", 0x10)
            else:
                print(message, file=sys.stderr)
        except Exception:  # noqa: BLE001
            pass


def main() -> int:
    log = app_logger.setup_logging()
    _install_exception_hooks(log)
    _enable_dpi_awareness()

    if os.name != "nt":
        log.error("Platform tidak didukung: %s", sys.platform)
        _show_fatal_dialog("UnsupportedPlatform")
        return 2

    log.info("WordSplitter dimulai. Python %s", sys.version.split()[0])
    try:
        import gui

        gui.run()
    except Exception:  # noqa: BLE001
        log.exception("Kegagalan fatal pada level aplikasi.")
        _show_fatal_dialog("ApplicationStartupError")
        return 1
    log.info("WordSplitter selesai.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

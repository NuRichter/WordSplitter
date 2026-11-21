"""Render the window (and optionally the guide) for visual QA (developer tool).

Usage: py -3 tools\\screenshot.py <out.png> [language] [light|dark] [guide|sample|analyzed]
The capture uses PrintWindow, so other windows on top do not spoil it.
"""

import ctypes
import os
import sys
from ctypes import wintypes as W
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))


def grab(widget, path: str) -> None:
    from PIL import Image

    user32, gdi32 = ctypes.windll.user32, ctypes.windll.gdi32
    hwnd = user32.GetParent(widget.winfo_id()) or widget.winfo_id()
    rect = W.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    width, height = rect.right - rect.left, rect.bottom - rect.top
    hdc = user32.GetWindowDC(hwnd)
    mem = gdi32.CreateCompatibleDC(hdc)
    bitmap = gdi32.CreateCompatibleBitmap(hdc, width, height)
    gdi32.SelectObject(mem, bitmap)
    user32.PrintWindow(hwnd, mem, 2)

    class BIH(ctypes.Structure):
        _fields_ = [("biSize", W.DWORD), ("biWidth", W.LONG), ("biHeight", W.LONG), ("biPlanes", W.WORD),
                    ("biBitCount", W.WORD), ("biCompression", W.DWORD), ("biSizeImage", W.DWORD),
                    ("x", W.LONG), ("y", W.LONG), ("u", W.DWORD), ("i", W.DWORD)]

    info = BIH(ctypes.sizeof(BIH), width, -height, 1, 32, 0, 0, 0, 0, 0, 0)
    buffer = ctypes.create_string_buffer(width * height * 4)
    gdi32.GetDIBits(mem, bitmap, 0, height, buffer, ctypes.byref(info), 0)
    Image.frombuffer("RGBA", (width, height), buffer, "raw", "BGRA", 0, 1).convert("RGB").save(path)
    gdi32.DeleteObject(bitmap)
    gdi32.DeleteDC(mem)
    user32.ReleaseDC(hwnd, hdc)


def main() -> None:
    out = sys.argv[1]
    language = sys.argv[2] if len(sys.argv) > 2 else "en"
    theme_name = sys.argv[3] if len(sys.argv) > 3 else "light"
    mode = sys.argv[4] if len(sys.argv) > 4 else ""
    import settings

    # Use a private settings file so the developer's own preferences stay untouched.
    os.environ["LOCALAPPDATA"] = str(Path(os.environ.get("TEMP", ".")) / "ws-shot")
    settings.Settings(language=language, theme=theme_name).save()
    import winsys

    winsys.enable_dpi_awareness()
    import gui

    scaling = None
    if os.environ.get("WS_SMALL"):
        # Emulate a 1366x768 laptop at 100 % scaling (96 DPI, 40 px taskbar).
        winsys.work_area = lambda: (0, 0, 1366, 728)
        scaling = 96 / 72
    root = gui.build_root(scaling)
    app = root._ws_app

    def stage() -> None:
        if mode == "analyzed":
            app.source = Path(r"C:\Users\Example\Documents\Skripsi Bab 1-5.docx")
            app.total_pages = 48
            app.fingerprint = [1, 1]
            app.file_status = "file.status.ready"
            app.var_page_a.set("24")
            app.var_page_b.set("25")
            app.status = ("status.analyzed", {})
            app._refresh_all()
        root.update()
        if mode == "guide":
            app.open_guide()
            root.update()
            root.after(1500, lambda: (grab(root.winfo_children()[-1], out), root.destroy()))
        else:
            root.after(400, lambda: (grab(root, out), root.destroy()))

    root.after(300, stage)
    root.mainloop()


if __name__ == "__main__":
    main()

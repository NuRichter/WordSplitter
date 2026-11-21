"""Thin ctypes layer over the Win32 calls WordSplitter needs.

Everything here degrades to a harmless no-op on other platforms or when an API
is missing, so callers never need their own platform checks. Nothing in this
module imports pywin32, which keeps the GUI process free of COM entirely.
"""

from __future__ import annotations

import ctypes
import os
import sys
from typing import List, Optional, Tuple

IS_WINDOWS = os.name == "nt"

if IS_WINDOWS:
    from ctypes import wintypes as W

    _user32 = ctypes.WinDLL("user32", use_last_error=True)
    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _kernel32.OpenProcess.restype = W.HANDLE
    _kernel32.OpenProcess.argtypes = (W.DWORD, W.BOOL, W.DWORD)
    _kernel32.CreateJobObjectW.restype = W.HANDLE
    _kernel32.CreateJobObjectW.argtypes = (ctypes.c_void_p, W.LPCWSTR)
    _kernel32.GetCurrentProcess.restype = W.HANDLE
    _kernel32.GetStdHandle.restype = W.HANDLE
    _kernel32.CreateToolhelp32Snapshot.restype = W.HANDLE
    _kernel32.CreateToolhelp32Snapshot.argtypes = (W.DWORD, W.DWORD)
    _user32.FindWindowW.restype = W.HWND
    _user32.FindWindowW.argtypes = (W.LPCWSTR, W.LPCWSTR)
    _user32.GetParent.restype = W.HWND
    _user32.GetParent.argtypes = (W.HWND,)

PROCESS_TERMINATE = 0x0001
PROCESS_SET_QUOTA = 0x0100
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
SYNCHRONIZE = 0x00100000
JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x2000
JOB_OBJECT_EXTENDED_LIMIT_INFORMATION_CLASS = 9


if IS_WINDOWS:

    class _BasicLimit(ctypes.Structure):
        _fields_ = [
            ("PerProcessUserTimeLimit", ctypes.c_int64),
            ("PerJobUserTimeLimit", ctypes.c_int64),
            ("LimitFlags", W.DWORD),
            ("MinimumWorkingSetSize", ctypes.c_size_t),
            ("MaximumWorkingSetSize", ctypes.c_size_t),
            ("ActiveProcessLimit", W.DWORD),
            ("Affinity", ctypes.c_size_t),
            ("PriorityClass", W.DWORD),
            ("SchedulingClass", W.DWORD),
        ]

    class _IoCounters(ctypes.Structure):
        _fields_ = [(name, ctypes.c_uint64) for name in ("r", "w", "o", "rb", "wb", "ob")]

    class _ExtendedLimit(ctypes.Structure):
        _fields_ = [
            ("BasicLimitInformation", _BasicLimit),
            ("IoInfo", _IoCounters),
            ("ProcessMemoryLimit", ctypes.c_size_t),
            ("JobMemoryLimit", ctypes.c_size_t),
            ("PeakProcessMemoryUsed", ctypes.c_size_t),
            ("PeakJobMemoryUsed", ctypes.c_size_t),
        ]


class KillOnCloseJob:
    """A Windows job object that terminates its members when its handle closes.

    The handle is owned by this process and is never inherited, so the kernel
    closes it when this process exits for any reason, including a crash or a
    forced termination from Task Manager. Every process assigned to the job is
    then terminated as well. This is what guarantees that no invisible
    WINWORD.EXE or worker process outlives WordSplitter.
    """

    def __init__(self) -> None:
        self._handle = None
        if not IS_WINDOWS:
            return
        handle = _kernel32.CreateJobObjectW(None, None)
        if not handle:
            return
        info = _ExtendedLimit()
        info.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        ok = _kernel32.SetInformationJobObject(
            W.HANDLE(handle),
            JOB_OBJECT_EXTENDED_LIMIT_INFORMATION_CLASS,
            ctypes.byref(info),
            ctypes.sizeof(info),
        )
        if ok:
            self._handle = handle
        else:
            _kernel32.CloseHandle(W.HANDLE(handle))

    @property
    def available(self) -> bool:
        return self._handle is not None

    def assign_pid(self, pid: int) -> bool:
        if not self._handle or not pid:
            return False
        process = _kernel32.OpenProcess(PROCESS_SET_QUOTA | PROCESS_TERMINATE, False, pid)
        if not process:
            return False
        try:
            return bool(_kernel32.AssignProcessToJobObject(W.HANDLE(self._handle), W.HANDLE(process)))
        finally:
            _kernel32.CloseHandle(W.HANDLE(process))

    def assign_handle(self, process_handle: int) -> bool:
        if not self._handle or not process_handle:
            return False
        return bool(
            _kernel32.AssignProcessToJobObject(W.HANDLE(self._handle), W.HANDLE(process_handle))
        )


def terminate_pid(pid: Optional[int]) -> bool:
    """Forcefully end one process. Returns True when the call was accepted."""
    if not IS_WINDOWS or not pid:
        return False
    process = _kernel32.OpenProcess(PROCESS_TERMINATE, False, pid)
    if not process:
        return False
    try:
        return bool(_kernel32.TerminateProcess(W.HANDLE(process), 1))
    finally:
        _kernel32.CloseHandle(W.HANDLE(process))


def pid_is_alive(pid: Optional[int]) -> bool:
    if not IS_WINDOWS or not pid:
        return False
    process = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION | SYNCHRONIZE, False, pid)
    if not process:
        return False
    try:
        return _kernel32.WaitForSingleObject(W.HANDLE(process), 0) == 0x102  # WAIT_TIMEOUT
    finally:
        _kernel32.CloseHandle(W.HANDLE(process))


def std_handle(which: int) -> Optional[int]:
    """A real standard handle (-10 input, -11 output), even in a windowed build."""
    if not IS_WINDOWS:
        return None
    handle = _kernel32.GetStdHandle(which)
    if not handle or handle == ctypes.c_void_p(-1).value:
        return None
    return int(handle)


def std_output_handle() -> Optional[int]:
    return std_handle(-11)


def std_input_handle() -> Optional[int]:
    return std_handle(-10)


def pid_of_window(class_name: str, title: str) -> Optional[int]:
    """Return the process id that owns a top level window, or None."""
    if not IS_WINDOWS:
        return None
    hwnd = _user32.FindWindowW(class_name, title)
    if not hwnd:
        return None
    pid = W.DWORD()
    _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value or None


def pids_by_image(image_name: str) -> set:
    """Return the ids of every running process with the given executable name."""
    found: set = set()
    if not IS_WINDOWS:
        return found

    class _PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [
            ("dwSize", W.DWORD), ("cntUsage", W.DWORD), ("th32ProcessID", W.DWORD),
            ("th32DefaultHeapID", ctypes.c_size_t), ("th32ModuleID", W.DWORD),
            ("cntThreads", W.DWORD), ("th32ParentProcessID", W.DWORD),
            ("pcPriClassBase", ctypes.c_long), ("dwFlags", W.DWORD), ("szExeFile", W.WCHAR * 260),
        ]

    snapshot = _kernel32.CreateToolhelp32Snapshot(0x2, 0)  # TH32CS_SNAPPROCESS
    if snapshot in (0, -1, ctypes.c_void_p(-1).value):
        return found
    try:
        entry = _PROCESSENTRY32W()
        entry.dwSize = ctypes.sizeof(entry)
        ok = _kernel32.Process32FirstW(W.HANDLE(snapshot), ctypes.byref(entry))
        target = image_name.lower()
        while ok:
            if entry.szExeFile.lower() == target:
                found.add(int(entry.th32ProcessID))
            ok = _kernel32.Process32NextW(W.HANDLE(snapshot), ctypes.byref(entry))
    finally:
        _kernel32.CloseHandle(W.HANDLE(snapshot))
    return found


def visible_windows_of(pid: int, with_handles: bool = False) -> list:
    """List (class, title) of every visible top level window owned by pid.

    A hidden automation instance of Word shows no dialog at all, so a visible
    dialog returned here is blocking the automation. With with_handles=True
    the tuples are (hwnd, class, title).
    """
    if not IS_WINDOWS or not pid:
        return []
    found: list = []

    @ctypes.WINFUNCTYPE(W.BOOL, W.HWND, W.LPARAM)
    def _callback(hwnd, _lparam):
        owner = W.DWORD()
        _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
        if owner.value == pid and _user32.IsWindowVisible(hwnd):
            title = ctypes.create_unicode_buffer(256)
            _user32.GetWindowTextW(hwnd, title, 256)
            klass = ctypes.create_unicode_buffer(128)
            _user32.GetClassNameW(hwnd, klass, 128)
            found.append((hwnd, klass.value, title.value) if with_handles else (klass.value, title.value))
        return True

    _user32.EnumWindows(_callback, 0)
    return found


def hide_window(hwnd) -> None:
    """Hide a window of another process without touching its message loop."""
    if IS_WINDOWS and hwnd:
        try:
            _user32.ShowWindowAsync(W.HWND(hwnd), 0)  # SW_HIDE
        except Exception:  # noqa: BLE001
            pass


def user_ui_language_tag() -> Optional[str]:
    """Return the Windows display language as a tag such as 'id-ID'."""
    if not IS_WINDOWS:
        return None
    try:
        buffer = ctypes.create_unicode_buffer(85)
        langid = _kernel32.GetUserDefaultUILanguage()
        # LCIDToLocaleName gives a BCP 47 style tag for the language id.
        if _kernel32.LCIDToLocaleName(langid, buffer, 85, 0):
            return buffer.value
    except Exception:  # noqa: BLE001
        pass
    return None


def system_prefers_dark() -> bool:
    if not IS_WINDOWS:
        return False
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        ) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return int(value) == 0
    except Exception:  # noqa: BLE001
        return False


def set_dark_title_bar(tk_window, dark: bool) -> None:
    """Ask DWM to paint the native title bar dark or light (Windows 10 1809+)."""
    if not IS_WINDOWS:
        return
    try:
        hwnd = _user32.GetParent(tk_window.winfo_id()) or tk_window.winfo_id()
        value = ctypes.c_int(1 if dark else 0)
        dwm = ctypes.WinDLL("dwmapi")
        # Attribute 20 is DWMWA_USE_IMMERSIVE_DARK_MODE; builds before 19041 used 19.
        if dwm.DwmSetWindowAttribute(W.HWND(hwnd), 20, ctypes.byref(value), ctypes.sizeof(value)) != 0:
            dwm.DwmSetWindowAttribute(W.HWND(hwnd), 19, ctypes.byref(value), ctypes.sizeof(value))
        # Force a non client repaint so the change is visible immediately.
        _user32.SetWindowPos(W.HWND(hwnd), None, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0004 | 0x0020)
    except Exception:  # noqa: BLE001 - purely cosmetic
        pass


def work_area() -> Optional[Tuple[int, int, int, int]]:
    """Return the primary monitor work area (left, top, right, bottom) in pixels."""
    if not IS_WINDOWS:
        return None
    try:
        rect = W.RECT()
        if _user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0):  # SPI_GETWORKAREA
            return rect.left, rect.top, rect.right, rect.bottom
    except Exception:  # noqa: BLE001
        pass
    return None


def trim_working_set() -> None:
    """Return unused pages to the OS. Windows does this anyway under pressure."""
    if not IS_WINDOWS:
        return
    try:
        _kernel32.SetProcessWorkingSetSize(_kernel32.GetCurrentProcess(), ctypes.c_size_t(-1), ctypes.c_size_t(-1))
    except Exception:  # noqa: BLE001
        pass


def enable_dpi_awareness() -> None:
    if not IS_WINDOWS:
        return
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system aware, matches Tk scaling
    except Exception:  # noqa: BLE001
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:  # noqa: BLE001
            pass


def drive_has_recycle_bin(path: str) -> bool:
    """Only local fixed drives keep a Recycle Bin; network and USB drives do not."""
    if not IS_WINDOWS:
        return False
    try:
        root = os.path.splitdrive(os.path.abspath(path))[0]
        if not root or root.startswith("\\\\"):
            return False
        return ctypes.windll.kernel32.GetDriveTypeW(root + "\\") == 3  # DRIVE_FIXED
    except Exception:  # noqa: BLE001
        return False


def move_to_recycle_bin(path: str) -> bool:
    """Send a file to the Recycle Bin without any UI. Returns True on success.

    Callers must check drive_has_recycle_bin first: on drives without a bin
    the shell would delete permanently, which is a different promise.
    """
    if not IS_WINDOWS:
        return False

    class _SHFILEOPSTRUCTW(ctypes.Structure):
        _fields_ = [
            ("hwnd", W.HWND),
            ("wFunc", ctypes.c_uint),
            ("pFrom", W.LPCWSTR),
            ("pTo", W.LPCWSTR),
            ("fFlags", ctypes.c_ushort),
            ("fAnyOperationsAborted", W.BOOL),
            ("hNameMappings", ctypes.c_void_p),
            ("lpszProgressTitle", W.LPCWSTR),
        ]

    FO_DELETE = 3
    FOF_SILENT, FOF_NOCONFIRMATION, FOF_ALLOWUNDO, FOF_NOERRORUI = 0x4, 0x10, 0x40, 0x400
    operation = _SHFILEOPSTRUCTW()
    operation.wFunc = FO_DELETE
    operation.pFrom = os.path.abspath(path) + "\0"  # the API needs a double NUL terminator
    operation.fFlags = FOF_SILENT | FOF_NOCONFIRMATION | FOF_ALLOWUNDO | FOF_NOERRORUI
    try:
        result = ctypes.windll.shell32.SHFileOperationW(ctypes.byref(operation))
    except Exception:  # noqa: BLE001
        return False
    return result == 0 and not operation.fAnyOperationsAborted and not os.path.exists(path)


def documents_dir() -> Optional[str]:
    """The user's real Documents folder, including OneDrive redirection."""
    if not IS_WINDOWS:
        return None
    try:
        class _GUID(ctypes.Structure):
            _fields_ = [("Data1", W.DWORD), ("Data2", W.WORD), ("Data3", W.WORD), ("Data4", ctypes.c_ubyte * 8)]

        # FOLDERID_Documents {FDD39AD0-238F-46AF-ADB4-6C85480369C7}
        folder = _GUID(0xFDD39AD0, 0x238F, 0x46AF, (ctypes.c_ubyte * 8)(0xAD, 0xB4, 0x6C, 0x85, 0x48, 0x03, 0x69, 0xC7))
        result = ctypes.c_wchar_p()
        if ctypes.windll.shell32.SHGetKnownFolderPath(ctypes.byref(folder), 0, None, ctypes.byref(result)) == 0:
            try:
                return result.value
            finally:
                ctypes.windll.ole32.CoTaskMemFree(result)
    except Exception:  # noqa: BLE001
        pass
    return None


def message_box(text: str, title: str, error: bool = True) -> None:
    """Last resort dialog that needs neither Tk nor COM."""
    if not IS_WINDOWS:
        print(f"{title}: {text}", file=sys.stderr)
        return
    try:
        _user32.MessageBoxW(None, text, title, 0x10 if error else 0x40)
    except Exception:  # noqa: BLE001
        pass

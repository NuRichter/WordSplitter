"""Microsoft Word COM automation layer (worker process only).

Rationale
---------
A .docx package stores a flow of content, not a sequence of pages. Pagination
is produced by Word's layout engine and depends on fonts, the printer driver
and compatibility options of the machine, so Word itself is the only reliable
authority for "where does page N start". This module wraps that engine and
adds the guard rails that unattended automation needs:

* Late bound (dynamic) dispatch only. The win32com gen_py cache is never
  touched, so a corrupted cache on the user's machine cannot break the app.
* Every COM call goes through a proxy that retries the transient
  "call rejected / retry later" answers Word gives while it is busy.
* The Word process id is resolved through a unique window caption. The
  process is placed in a kill-on-close job object and watched by a watchdog:
  a hidden automation instance never owns a visible window, so any visible
  window is a modal dialog that would block forever. The watchdog records its
  title and terminates Word, turning a hang into a clear error.
* Word options that persist in the user's profile are restored before quit.
"""

from __future__ import annotations

import logging
import threading
import time
import uuid
from pathlib import Path
from typing import Any, Callable, List, Optional, Tuple

import winsys
from errors import AppError, OperationCancelled

LOGGER = logging.getLogger("WordSplitter.word")

WORD_CLSID = "{000209FF-0000-0000-C000-000000000046}"

# Word enumeration constants, declared locally so that no type library is needed.
WD_STATISTIC_PAGES = 2
WD_GOTO_PAGE = 1
WD_GOTO_ABSOLUTE = 1
WD_DO_NOT_SAVE_CHANGES = 0
WD_FORMAT_DOCUMENT_DEFAULT = 16  # .docx
WD_ALERTS_NONE = 0
WD_NO_PROTECTION = -1
WD_WITHIN_TABLE = 12
WD_LINE_SPACE_EXACTLY = 4
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
BREAK_CHARS = ("\x0c", "\x0e")  # page or section break, column break
NO_PASSWORD = "#wordsplitter-no-password#"

HR_CALL_REJECTED = -2147418111      # RPC_E_CALL_REJECTED
HR_RETRY_LATER = -2147417846        # RPC_E_SERVERCALL_RETRYLATER
TRANSIENT_HRESULTS = {HR_CALL_REJECTED, HR_RETRY_LATER}
DEAD_HRESULTS = {
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2147023170,  # RPC_S_CALL_FAILED
    -2147023169,  # RPC_S_CALL_FAILED_DNE
    -2147417848,  # RPC_E_DISCONNECTED
    -2147417856,  # RPC_E_SYS_CALL_FAILED
    -2147418094,  # RPC_E_SERVERFAULT
}
TRANSIENT_RETRY_SECONDS = 120.0

PAGESETUP_PROPERTIES = (
    "Orientation", "PageWidth", "PageHeight", "TopMargin", "BottomMargin", "LeftMargin",
    "RightMargin", "Gutter", "GutterPos", "HeaderDistance", "FooterDistance", "MirrorMargins",
    "VerticalAlignment", "DifferentFirstPageHeaderFooter", "OddAndEvenPagesHeaderFooter",
    "LayoutMode", "CharsLine", "LinesPage", "SectionDirection", "TwoPagesOnOne", "BookFoldPrinting",
)
PAGE_NUMBER_PROPERTIES = (
    "NumberStyle", "IncludeChapterNumber", "HeadingLevelForChapter", "ChapterPageSeparator",
    "RestartNumberingAtSection", "StartingNumber",
)
PERSISTENT_OPTIONS = (
    # (option, value used during automation). Originals are restored on quit.
    ("WarnBeforeSavingPrintingSendingMarkup", False),
    ("UpdateLinksAtOpen", False),
    ("BackgroundSave", False),
)


_PENDING_EXITS: set = set()
# Set once the worker has been told to stop. Sessions created afterwards see it
# immediately, so a cancel that races with Word start up is never lost.
CANCEL_ALL = threading.Event()


def settle_exits(timeout: float = 10.0) -> None:
    """Wait for quitting Word processes to end; terminate any that hang."""
    deadline = time.monotonic() + timeout
    while _PENDING_EXITS:
        for pid in list(_PENDING_EXITS):
            if not winsys.pid_is_alive(pid):
                _PENDING_EXITS.discard(pid)
        if not _PENDING_EXITS:
            return
        if time.monotonic() > deadline:
            for pid in list(_PENDING_EXITS):
                LOGGER.warning("Word pid %s did not exit after Quit; terminating it.", pid)
                winsys.terminate_pid(pid)
            _PENDING_EXITS.clear()
            return
        time.sleep(0.05)


class WordGone(Exception):
    """The Word process ended while a call was in flight."""


def _hresult(exc: BaseException) -> Optional[int]:
    args = getattr(exc, "args", ())
    if args and isinstance(args[0], int):
        return args[0]
    return None


def com_message(exc: BaseException) -> str:
    """Extract a readable message from a COM exception."""
    args = getattr(exc, "args", ())
    if len(args) > 2 and isinstance(args[2], tuple) and len(args[2]) > 2 and args[2][2]:
        return str(args[2][2]).strip()
    if len(args) > 1 and args[1]:
        return str(args[1])
    return str(exc) or type(exc).__name__


class _Com:
    """Transparent proxy that routes every COM access through Session.call."""

    __slots__ = ("_obj", "_session")

    def __init__(self, obj: Any, session: "WordSession") -> None:
        object.__setattr__(self, "_obj", obj)
        object.__setattr__(self, "_session", session)

    def __getattr__(self, name: str) -> Any:
        if name.startswith("__"):
            raise AttributeError(name)
        value = self._session.call(getattr, self._obj, name)
        return self._session.wrap(value)

    def __setattr__(self, name: str, value: Any) -> None:
        self._session.call(setattr, self._obj, name, _unwrap(value))

    def __call__(self, *args: Any, **kwargs: Any) -> Any:
        args = tuple(_unwrap(a) for a in args)
        kwargs = {k: _unwrap(v) for k, v in kwargs.items()}
        return self._session.wrap(self._session.call(self._obj, *args, **kwargs))


def _unwrap(value: Any) -> Any:
    return object.__getattribute__(value, "_obj") if isinstance(value, _Com) else value


class WordSession:
    """A dedicated, invisible and supervised Microsoft Word instance."""

    WATCH_INTERVAL = 0.35
    WATCH_STRIKES = 4  # about 1.4 s of a persistent visible window
    active: "set" = set()  # sessions alive in this process, for cancel_all()

    def __init__(self, job: Optional[winsys.KillOnCloseJob] = None) -> None:
        self._pythoncom: Any = None
        self._raw_app: Any = None
        self.app: Any = None
        self.pid: Optional[int] = None
        self._job = job
        self._watch_stop = threading.Event()
        self._watch_thread: Optional[threading.Thread] = None
        self.blocked_by: Optional[str] = None
        self._saved_options: List[Tuple[str, Any]] = []
        self._cancelled = False

    # -- lifecycle ---------------------------------------------------------

    def __enter__(self) -> "WordSession":
        self.start()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def start(self) -> None:
        try:
            import pythoncom  # type: ignore
            import pywintypes  # type: ignore
            from win32com.client import dynamic  # type: ignore
        except ImportError as exc:  # pragma: no cover - Windows only path
            raise AppError("err.word_missing", detail=f"pywin32: {exc}") from exc

        if self.cancelled:
            raise OperationCancelled()
        self._pythoncom = pythoncom
        pythoncom.CoInitialize()
        # Registered before Word exists, so a cancel that arrives while Word is
        # still starting is not lost: start() sees the flag as soon as it can.
        WordSession.active.add(self)
        before = winsys.pids_by_image("WINWORD.EXE")
        try:
            unknown = pythoncom.CoCreateInstanceEx(
                pywintypes.IID(WORD_CLSID), None, pythoncom.CLSCTX_LOCAL_SERVER, None,
                (pythoncom.IID_IDispatch,),
            )[0]
            self._raw_app = dynamic.Dispatch(unknown)
        except Exception as exc:  # noqa: BLE001
            WordSession.active.discard(self)
            self._uninitialize()
            if self.cancelled:
                raise OperationCancelled() from exc
            raise AppError("err.word_missing", detail=com_message(exc)) from exc

        self.app = _Com(self._raw_app, self)
        try:
            # First guess: the single WINWORD.EXE that did not exist a moment
            # ago. It lets the watchdog cover even the very first call.
            fresh = winsys.pids_by_image("WINWORD.EXE") - before
            if len(fresh) == 1:
                self._adopt_pid(fresh.pop())
            # Definitive answer: the process owning our uniquely captioned window.
            self._adopt_pid(self._pid_from_caption())
            if not self.pid:
                LOGGER.warning("Word pid unknown; dialog watchdog disabled.")
            if self.cancelled:
                raise OperationCancelled()
            self._configure()
        except BaseException as exc:
            failure = self.failure(exc, "err.word_missing")
            self.close()
            raise failure from exc
        LOGGER.info("Word started (pid %s, version %s).", self.pid, self._safe_version())

    def _adopt_pid(self, pid: Optional[int]) -> None:
        if not pid or pid == self.pid:
            return
        self.pid = pid
        if self._job is not None and self._job.available and not self._job.assign_pid(pid):
            LOGGER.warning("Word pid %s could not join the kill-on-close job.", pid)
        if self._watch_thread is None:
            self._watch_thread = threading.Thread(target=self._watch, name="WordWatchdog", daemon=True)
            self._watch_thread.start()

    def _pid_from_caption(self) -> Optional[int]:
        caption = f"WordSplitter-{uuid.uuid4().hex}"
        try:
            self.app.Caption = caption
        except WordGone:
            raise
        except Exception as exc:  # noqa: BLE001
            LOGGER.debug("Caption based pid lookup failed: %s", com_message(exc))
            return None
        for _ in range(60):
            pid = winsys.pid_of_window("OpusApp", caption)
            if pid:
                return pid
            time.sleep(0.02)
        return None

    def _safe_version(self) -> str:
        try:
            return str(self.app.Version)
        except Exception:  # noqa: BLE001
            return "?"

    def _configure(self) -> None:
        """Silence every interactive behaviour that could block automation."""
        for attribute, value in (
            ("Visible", False),
            ("DisplayAlerts", WD_ALERTS_NONE),
            ("ScreenUpdating", False),
            ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
        ):
            try:
                setattr(self.app, attribute, value)
            except WordGone:
                raise
            except Exception as exc:  # noqa: BLE001
                LOGGER.debug("Application.%s not set: %s", attribute, com_message(exc))
        options = None
        try:
            options = self.app.Options
        except WordGone:
            raise
        except Exception:  # noqa: BLE001
            pass
        if options is None:
            return
        for name, value in PERSISTENT_OPTIONS:
            try:
                current = getattr(options, name)
                if current != value:
                    setattr(options, name, value)
                    self._saved_options.append((name, current))
            except WordGone:
                raise
            except Exception as exc:  # noqa: BLE001
                LOGGER.debug("Options.%s not set: %s", name, com_message(exc))

    def _restore_options(self) -> None:
        if not self._saved_options:
            return
        try:
            options = self.app.Options
            for name, value in self._saved_options:
                try:
                    setattr(options, name, value)
                except Exception:  # noqa: BLE001
                    pass
        except Exception:  # noqa: BLE001
            pass
        self._saved_options = []

    def close(self) -> None:
        self._watch_stop.set()
        WordSession.active.discard(self)
        if (self.cancelled or self.blocked_by is not None) and self.pid:
            winsys.terminate_pid(self.pid)
        raw = self._raw_app
        if raw is not None:
            alive = self.pid is None or winsys.pid_is_alive(self.pid)
            if alive and self.blocked_by is None:
                if not self.cancelled:
                    self._restore_options()
                    try:
                        self.app.NormalTemplate.Saved = True
                    except BaseException:  # noqa: BLE001
                        pass
                try:
                    # The raw object, not the proxy: the proxy refuses every
                    # call after a cancel, but an unknown pid leaves Quit as
                    # the only way to stop this instance.
                    raw.Quit(WD_DO_NOT_SAVE_CHANGES)
                except Exception as exc:  # noqa: BLE001
                    LOGGER.debug("Word did not quit normally: %s", com_message(exc))
            self._raw_app = None
            self.app = None
        if self.pid and winsys.pid_is_alive(self.pid):
            if self.blocked_by is None and not self.cancelled:
                # Quit returns at once but WINWORD.EXE needs a couple of seconds
                # to shut down. The worker reports its result first and then
                # waits for this in settle_exits(), so nobody waits on it.
                _PENDING_EXITS.add(self.pid)
            else:
                winsys.terminate_pid(self.pid)
        self._uninitialize()
        LOGGER.info("Word session closed.")

    def _uninitialize(self) -> None:
        if self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass
            self._pythoncom = None

    @property
    def cancelled(self) -> bool:
        return self._cancelled or CANCEL_ALL.is_set()

    def kill(self) -> None:
        if self.pid:
            winsys.terminate_pid(self.pid)

    def cancel(self) -> None:
        """Called from another thread: stop Word now, whatever it is doing."""
        self._cancelled = True
        self.kill()

    @classmethod
    def cancel_all(cls, wait_for_pid: float = 30.0) -> None:
        """Cancel every live session, including one still starting Word.

        While CoCreateInstanceEx is running the Word pid is not known yet, so
        the pid cannot be terminated. The session is flagged instead; start()
        checks the flag the moment Word is up and close() terminates it. The
        loop below keeps killing until every flagged session knows its pid.
        """
        CANCEL_ALL.set()
        deadline = time.monotonic() + wait_for_pid
        while True:
            pending = False
            for session in list(cls.active):
                session.cancel()
                if not session.pid:
                    pending = True
            if not pending or time.monotonic() > deadline:
                return
            time.sleep(0.1)

    # -- supervision -------------------------------------------------------

    def _watch(self) -> None:
        strikes = 0
        while not self._watch_stop.wait(self.WATCH_INTERVAL):
            windows = []
            for hwnd, klass, title in winsys.visible_windows_of(self.pid or 0, with_handles=True):
                if klass == "OpusApp":
                    # Word sometimes surfaces its own document window, for
                    # example when a document is activated. That is not a
                    # question waiting for an answer: hide it and carry on.
                    winsys.hide_window(hwnd)
                    continue
                windows.append((klass, title))
            if not windows:
                strikes = 0
                continue
            strikes += 1
            if strikes >= self.WATCH_STRIKES:
                title = next((t for _c, t in windows if t), windows[0][0])
                self.blocked_by = title or "Microsoft Word"
                LOGGER.error("Word blocked by a dialog %r (%s); terminating it.", title, windows)
                winsys.terminate_pid(self.pid)
                return

    # -- call routing ------------------------------------------------------

    def wrap(self, value: Any) -> Any:
        if value is None or isinstance(value, (bool, int, float, str, bytes, tuple)):
            return value
        if isinstance(value, _Com):
            return value
        return _Com(value, self)

    def call(self, fn: Callable, *args: Any, **kwargs: Any) -> Any:
        deadline = None
        delay = 0.05
        while True:
            if self.cancelled:
                raise OperationCancelled()
            if self.blocked_by is not None:
                raise WordGone(self.blocked_by)
            try:
                return fn(*args, **kwargs)
            except Exception as exc:  # noqa: BLE001
                code = _hresult(exc)
                if code in TRANSIENT_HRESULTS:
                    now = time.monotonic()
                    deadline = deadline or now + TRANSIENT_RETRY_SECONDS
                    if now < deadline:
                        time.sleep(delay)
                        delay = min(delay * 2, 1.0)
                        continue
                if code in DEAD_HRESULTS or self.blocked_by is not None:
                    raise WordGone(self.blocked_by or com_message(exc)) from exc
                raise

    def failure(self, exc: BaseException, key: str) -> Exception:
        """Translate any exception raised inside a session into AppError or cancel."""
        if self.cancelled or isinstance(exc, OperationCancelled):
            return OperationCancelled()
        if isinstance(exc, AppError):
            return exc
        if self.blocked_by is not None:
            return AppError("err.word_blocked", title=self.blocked_by)
        if isinstance(exc, WordGone):
            return AppError("err.word_crashed", detail=str(exc))
        return AppError(key, detail=com_message(exc))

    # -- documents ---------------------------------------------------------

    def open(self, path: Path) -> Any:
        try:
            return self.app.Documents.Open(
                FileName=str(path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                PasswordDocument=NO_PASSWORD,
                PasswordTemplate=NO_PASSWORD,
                Revert=False,
                WritePasswordDocument=NO_PASSWORD,
                WritePasswordTemplate=NO_PASSWORD,
                Visible=False,
                OpenAndRepair=False,
                NoEncodingDialog=True,
            )
        except Exception as exc:  # noqa: BLE001
            text = com_message(exc).lower()
            if "password" in text:
                raise AppError("err.encrypted_or_legacy", detail=com_message(exc)) from exc
            raise self.failure(exc, "err.word_open") from exc

    def close_document(self, document: Any) -> None:
        if document is None:
            return
        try:
            document.Close(WD_DO_NOT_SAVE_CHANGES)
        except Exception as exc:  # noqa: BLE001
            LOGGER.debug("Document close failed: %s", com_message(exc))

    def page_count(self, document: Any) -> int:
        try:
            try:
                document.Repaginate()
            except WordGone:
                raise
            except Exception as exc:  # noqa: BLE001
                LOGGER.debug("Repaginate failed: %s", com_message(exc))
            pages = int(document.ComputeStatistics(WD_STATISTIC_PAGES))
        except Exception as exc:  # noqa: BLE001
            raise self.failure(exc, "err.word_pagination") from exc
        if pages < 1:
            raise AppError("err.word_pagination", detail=f"pages={pages}")
        return pages

    def page_start(self, document: Any, page: int) -> int:
        """Character offset at which a rendered page begins (Word's own layout)."""
        try:
            offset = int(document.GoTo(What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=page).Start)
            content = document.Content
            start, end = int(content.Start), int(content.End)
        except Exception as exc:  # noqa: BLE001
            raise self.failure(exc, "err.word_pagination") from exc
        if not start < offset < end:
            raise AppError("err.empty_part", detail=f"offset={offset} range={start}-{end}")
        return offset

    def save_docx(self, document: Any, target: Path) -> None:
        try:
            document.SaveAs2(FileName=str(target), FileFormat=WD_FORMAT_DOCUMENT_DEFAULT, AddToRecentFiles=False)
        except Exception as exc:  # noqa: BLE001
            raise self.failure(exc, "err.save_failed") from exc

    def cut_keep_before(self, document: Any, split: int) -> "CutReport":
        """Turn the open document into Part 1 (everything before `split`)."""
        try:
            return _cut_before(document, split)
        except AppError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise self.failure(exc, "err.cut_failed") from exc

    def cut_keep_from(self, document: Any, split: int) -> "CutReport":
        """Turn the open document into Part 2 (everything from `split` on)."""
        try:
            return _cut_from(document, split)
        except AppError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise self.failure(exc, "err.cut_failed") from exc

    def reopen_check(self, path: Path, label: str) -> None:
        """A produced file must be something Word itself opens again."""
        document = None
        try:
            try:
                document = self.open(path)
                empty = int(document.Content.End) < 1
            except AppError as exc:
                raise AppError("err.verify_failed", detail=exc.detail or exc.key, label=label) from exc
            except Exception as exc:  # noqa: BLE001
                error = self.failure(exc, "err.verify_failed")
                error.params.setdefault("label", label)
                raise error from exc
            if empty:
                raise AppError("err.verify_failed", label=label)
        finally:
            self.close_document(document)


# -- cutting --------------------------------------------------------------------


class CutReport:
    """Notes collected while one part is being cut."""

    def __init__(self) -> None:
        self.fields_unlinked = 0
        self.header_rows_kept = 0
        self.steps: List[str] = []


def _char_at(document: Any, position: int) -> str:
    return document.Range(position, position + 1).Text or ""


def _snap_to_row(document: Any, split: int, report: CutReport) -> int:
    """A page that starts inside a table row is cut at the start of that row."""
    probe = document.Range(split, split)
    try:
        if probe.Information(WD_WITHIN_TABLE):
            row_start = int(probe.Rows(1).Range.Start)
            if row_start < split:
                report.steps.append(f"snapped {split}->{row_start}")
                return row_start
    except WordGone:
        raise
    except Exception as exc:  # noqa: BLE001 - vertically merged cells have no rows
        report.steps.append(f"row snap skipped: {com_message(exc)}")
    return split


def _prepare(document: Any, report: CutReport) -> Tuple[bool, int]:
    """Make the working copy editable without tracking. Returns the old state."""
    track = False
    protection = WD_NO_PROTECTION
    try:
        track = bool(document.TrackRevisions)
        if track:
            document.TrackRevisions = False
    except WordGone:
        raise
    except Exception as exc:  # noqa: BLE001
        report.steps.append(f"track: {com_message(exc)}")
    try:
        if bool(document.Final):
            document.Final = False
    except WordGone:
        raise
    except Exception:  # noqa: BLE001
        pass
    try:
        protection = int(document.ProtectionType)
    except WordGone:
        raise
    except Exception:  # noqa: BLE001
        protection = WD_NO_PROTECTION
    if protection != WD_NO_PROTECTION:
        try:
            document.Unprotect()
        except WordGone:
            raise
        except Exception as exc:  # noqa: BLE001
            raise AppError("err.doc_restricted", detail=com_message(exc)) from exc
    return track, protection


def _restore(document: Any, state: Tuple[bool, int]) -> None:
    track, protection = state
    if protection != WD_NO_PROTECTION:
        try:
            document.Protect(Type=protection, NoReset=True)
        except WordGone:
            raise
        except Exception as exc:  # noqa: BLE001
            LOGGER.debug("Protection not restored: %s", com_message(exc))
    if track:
        try:
            document.TrackRevisions = True
        except WordGone:
            raise
        except Exception:  # noqa: BLE001
            pass


def _straddling_fields(document: Any, position: int) -> list:
    hits = []
    fields = document.Fields
    for index in range(1, int(fields.Count) + 1):
        field = fields(index)
        try:
            start = int(field.Code.Start) - 1
            result = field.Result
            end = (int(result.End) if result is not None else int(field.Code.End)) + 1
        except WordGone:
            raise
        except Exception:  # noqa: BLE001
            continue
        if start < position < end:
            hits.append(field)
    return hits


def _delete(document: Any, bounds: Callable[[], Tuple[int, int]], mark: Any, report: CutReport) -> None:
    """Delete a range; on refusal unlink fields and unlock controls at the cut.

    Word refuses to delete a range that begins or ends inside a field result
    (a table of contents, a Zotero or Mendeley bibliography) or inside a
    locked content control. Those objects are converted to static content at
    the cut only, then the deletion is retried. The live range `mark` tracks
    the cut position while those edits shift character offsets.
    """
    start, end = bounds()
    if end <= start:
        return
    try:
        document.Range(start, end).Delete()
        return
    except WordGone:
        raise
    except Exception as exc:  # noqa: BLE001
        report.steps.append(f"delete refused: {com_message(exc)}")

    for field in _straddling_fields(document, int(mark.Start)):
        try:
            field.Unlink()
            report.fields_unlinked += 1
        except WordGone:
            raise
        except Exception as exc:  # noqa: BLE001
            report.steps.append(f"unlink failed: {com_message(exc)}")
    controls = document.ContentControls
    for index in range(1, int(controls.Count) + 1):
        control = controls(index)
        try:
            span = control.Range
            if int(span.Start) - 1 <= int(mark.Start) <= int(span.End) + 1:
                control.LockContentControl = False
                control.LockContents = False
        except WordGone:
            raise
        except Exception:  # noqa: BLE001
            continue
    start, end = bounds()
    try:
        document.Range(start, end).Delete()
    except WordGone:
        raise
    except Exception as exc:  # noqa: BLE001
        raise AppError("err.cut_failed", detail=com_message(exc)) from exc


def _copy_story(source_hf: Any, target_hf: Any) -> None:
    """Copy one header or footer, dropping the extra paragraph Word appends."""
    target_hf.Range.FormattedText = source_hf.Range.FormattedText
    story = target_hf.Range
    if int(story.Paragraphs.Count) > 1:
        last = story.Paragraphs.Last.Range
        if (last.Text or "") in ("\r", ""):
            tail = story.Duplicate
            tail.SetRange(int(last.Start) - 1, int(last.Start))
            tail.Delete()


def _copy_section_format(document: Any, source_index: int, target_index: int, report: CutReport) -> None:
    """Give the last section the page setup, headers and numbering of another.

    The final paragraph mark carries the formatting of the last section. When
    Part 1 ends inside an earlier section, deleting the rest of the document
    would otherwise dress that content in the last section's layout.
    """
    source = document.Sections(source_index)
    target = document.Sections(target_index)
    source_setup, target_setup = source.PageSetup, target.PageSetup
    for name in PAGESETUP_PROPERTIES:
        try:
            value = getattr(source_setup, name)
            if getattr(target_setup, name) != value:
                setattr(target_setup, name, value)
        except WordGone:
            raise
        except Exception as exc:  # noqa: BLE001
            report.steps.append(f"pagesetup {name}: {com_message(exc)}")
    try:
        source_columns, target_columns = source_setup.TextColumns, target_setup.TextColumns
        if int(target_columns.Count) != int(source_columns.Count):
            target_columns.SetCount(int(source_columns.Count))
        if int(source_columns.Count) > 1:
            target_columns.EvenlySpaced = source_columns.EvenlySpaced
            target_columns.Spacing = source_columns.Spacing
    except WordGone:
        raise
    except Exception as exc:  # noqa: BLE001
        report.steps.append(f"columns: {com_message(exc)}")
    for collection in ("Headers", "Footers"):
        for kind in (1, 2, 3):  # primary, first page, even pages
            try:
                source_hf = getattr(source, collection)(kind)
                target_hf = getattr(target, collection)(kind)
                if target_hf.LinkToPrevious:
                    target_hf.LinkToPrevious = False
                _copy_story(source_hf, target_hf)
            except WordGone:
                raise
            except Exception as exc:  # noqa: BLE001
                report.steps.append(f"{collection}{kind}: {com_message(exc)}")
    try:
        source_numbers = source.Footers(1).PageNumbers
        target_numbers = target.Footers(1).PageNumbers
        for name in PAGE_NUMBER_PROPERTIES:
            try:
                setattr(target_numbers, name, getattr(source_numbers, name))
            except WordGone:
                raise
            except Exception:  # noqa: BLE001
                pass
    except WordGone:
        raise
    except Exception as exc:  # noqa: BLE001
        report.steps.append(f"page numbers: {com_message(exc)}")


def _tidy_tail(document: Any, report: CutReport) -> None:
    """Remove a trailing paragraph that holds only breaks (a blank last page)."""
    for _ in range(4):
        paragraphs = document.Paragraphs
        count = int(paragraphs.Count)
        if count < 2:
            return
        last = paragraphs.Last.Range
        text = last.Text or ""
        body = text[:-1] if text.endswith("\r") else text
        if body.strip("".join(BREAK_CHARS)) != "":
            return
        try:
            if int(last.Fields.Count) or int(last.InlineShapes.Count):
                return  # looks empty but holds a field end or a picture
        except WordGone:
            raise
        except Exception:  # noqa: BLE001
            pass
        last_start = int(last.Start)
        after_table = False
        try:
            after_table = bool(document.Range(last_start - 1, last_start - 1).Information(WD_WITHIN_TABLE))
        except WordGone:
            raise
        except Exception:  # noqa: BLE001
            pass
        # Everything below is cosmetic: if Word refuses, the paragraph stays.
        try:
            if after_table:
                # Word requires a paragraph after a table; make it as small as possible.
                last.Font.Size = 1
                fmt = last.ParagraphFormat
                fmt.SpaceBefore = 0
                fmt.SpaceAfter = 0
                fmt.LineSpacingRule = WD_LINE_SPACE_EXACTLY
                fmt.LineSpacing = 1
                fmt.PageBreakBefore = False
                report.steps.append("minimised paragraph after table")
                return
            document.Range(last_start, int(document.Content.End)).Delete()
            if int(document.Paragraphs.Count) >= count:
                document.Range(last_start - 1, last_start).Delete()
        except WordGone:
            raise
        except Exception as exc:  # noqa: BLE001
            report.steps.append(f"blank tail kept: {com_message(exc)}")
            return
        if int(document.Paragraphs.Count) >= count:
            report.steps.append("trailing paragraph kept")
            return
        report.steps.append("removed blank tail")


def _section_index(document: Any, position: int) -> int:
    return int(document.Range(position, position).Sections(1).Index)


def _cut_before(document: Any, split: int) -> CutReport:
    """Turn the open document into Part 1: everything before `split`."""
    report = CutReport()
    state = _prepare(document, report)
    cut = _snap_to_row(document, split, report)
    while cut > 1 and _char_at(document, cut - 1) in BREAK_CHARS:
        cut -= 1
    if cut <= int(document.Content.Start):
        raise AppError("err.empty_part")
    total_sections = int(document.Sections.Count)
    own_section = _section_index(document, cut - 1)
    if own_section < total_sections:
        _copy_section_format(document, own_section, total_sections, report)
        report.steps.append(f"section {own_section} formatting moved to {total_sections}")
    mark = document.Range(cut, cut)
    _delete(document, lambda: (int(mark.Start), int(document.Content.End)), mark, report)
    _tidy_tail(document, report)
    _restore(document, state)
    return report


def _drop_empty_lead_before_table(document: Any, report: CutReport) -> None:
    """Word leaves an empty paragraph in front of a table it was cut above."""
    try:
        first = document.Paragraphs(1).Range
        if (first.Text or "") != "\r" or first.Information(WD_WITHIN_TABLE):
            return
        if not document.Range(int(first.End), int(first.End)).Information(WD_WITHIN_TABLE):
            return
        tables = int(document.Tables.Count)
        first.Delete()
        if int(document.Tables.Count) == tables:
            report.steps.append("removed empty paragraph before table")
    except WordGone:
        raise
    except Exception as exc:  # noqa: BLE001 - cosmetic only
        report.steps.append(f"lead paragraph kept: {com_message(exc)}")


def _heading_rows(table: Any) -> int:
    count = 0
    rows = int(table.Rows.Count)
    while count < rows and bool(table.Rows(count + 1).HeadingFormat):
        count += 1
    return count


def _cut_from(document: Any, split: int) -> CutReport:
    """Turn the open document into Part 2: everything from `split` onwards."""
    report = CutReport()
    state = _prepare(document, report)
    start = _snap_to_row(document, split, report)
    end = int(document.Content.End)
    while start < end - 1 and _char_at(document, start) in BREAK_CHARS:
        start += 1
    if start >= end - 1:
        raise AppError("err.empty_part")

    own_section = _section_index(document, start)
    if own_section > 1:
        # Headers inherited from earlier sections would vanish with them;
        # unlinking makes Word copy their content into this section first.
        section = document.Sections(own_section)
        for collection in ("Headers", "Footers"):
            for kind in (1, 2, 3):
                try:
                    story = getattr(section, collection)(kind)
                    if story.LinkToPrevious:
                        story.LinkToPrevious = False
                except WordGone:
                    raise
                except Exception as exc:  # noqa: BLE001
                    report.steps.append(f"unlink {collection}{kind}: {com_message(exc)}")

    done = False
    probe = document.Range(start, start)
    try:
        if probe.Information(WD_WITHIN_TABLE):
            table = probe.Tables(1)
            row_index = int(probe.Rows(1).Index)
            heads = _heading_rows(table)
            if heads and row_index > heads + 1:
                # Keep the repeating header rows above the first kept row.
                first = int(table.Rows(heads + 1).Range.Start)
                last = int(table.Rows(row_index - 1).Range.End)
                document.Range(first, last).Rows.Delete()
                table_start = int(table.Range.Start)
                if table_start > 0:
                    mark = document.Range(table_start, table_start)
                    _delete(document, lambda: (0, int(mark.Start)), mark, report)
                report.header_rows_kept = heads
                done = True
                _drop_empty_lead_before_table(document, report)
    except WordGone:
        raise
    except AppError:
        raise
    except Exception as exc:  # noqa: BLE001
        report.steps.append(f"header rows not kept: {com_message(exc)}")
    if not done and start > 0:
        mark = document.Range(start, start)
        _delete(document, lambda: (0, int(mark.Start)), mark, report)
    _restore(document, state)
    return report

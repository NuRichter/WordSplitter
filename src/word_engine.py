"""Microsoft Word COM automation layer.

Rationale
---------
A .docx package stores a flow of content, not a sequence of pages. Pagination is
produced by the rendering engine at layout time and depends on the font metrics,
the printer driver and the compatibility options of the machine. Consequently
the only source of truth for the question "where does page N start" on Windows
11 is Microsoft Word itself. This module wraps that engine.

The split is performed by duplicating the source file on disk and deleting the
unwanted range inside each duplicate. The document is never rebuilt from
extracted text, therefore styles, sections, headers, footers, tables, images,
numbering and metadata survive the operation.
"""

from __future__ import annotations

import logging
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator, Optional

LOGGER = logging.getLogger("WordSplitter.word")

# Word enumeration constants, declared locally so that no type library is needed.
WD_STATISTIC_PAGES = 2
WD_GOTO_PAGE = 1
WD_GOTO_ABSOLUTE = 1
WD_DO_NOT_SAVE_CHANGES = 0
WD_FORMAT_DOCUMENT_DEFAULT = 16  # .docx
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


class WordEngineError(Exception):
    """Base class for every failure raised by the Word layer."""


class WordNotAvailableError(WordEngineError):
    """Microsoft Word desktop is not installed or cannot be started."""


class DocumentOpenError(WordEngineError):
    """The document could not be opened by Word."""


class PaginationError(WordEngineError):
    """Word could not produce a reliable pagination for the document."""


class SaveError(WordEngineError):
    """Word could not write the output document."""


def _import_com() -> tuple[Any, Any]:
    """Import the COM stack lazily so that this module imports on any platform."""
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except ImportError as exc:  # pragma: no cover - Windows only path
        raise WordNotAvailableError(
            "Komponen otomasi Windows (pywin32) tidak tersedia. "
            "Aplikasi ini hanya dapat berjalan pada Windows 11 dengan pywin32 terpasang."
        ) from exc
    return pythoncom, win32com.client


def _com_message(exc: BaseException) -> str:
    """Extract a human readable message from a COM exception."""
    detail = getattr(exc, "excepinfo", None)
    if isinstance(detail, tuple) and len(detail) > 2 and detail[2]:
        return str(detail[2]).strip()
    return str(exc)


class WordSession:
    """A dedicated, invisible Microsoft Word instance.

    A dedicated instance is created with DispatchEx so that the documents opened
    here never interfere with a Word window the user already has open, and so
    that quitting this session never closes the user's own work.
    """

    def __init__(self) -> None:
        self._pythoncom: Any = None
        self._client: Any = None
        self._app: Any = None

    # -- lifecycle ---------------------------------------------------------

    def __enter__(self) -> "WordSession":
        self._pythoncom, self._client = _import_com()
        try:
            self._pythoncom.CoInitialize()
        except Exception as exc:  # noqa: BLE001
            raise WordEngineError(
                f"Inisialisasi COM gagal: {_com_message(exc)}"
            ) from exc
        try:
            self._app = self._client.DispatchEx("Word.Application")
        except Exception as exc:  # noqa: BLE001
            self._safe_couninitialize()
            raise WordNotAvailableError(
                "Microsoft Word desktop tidak dapat dijalankan. "
                "Pastikan Microsoft Word (Office 2016 atau lebih baru, termasuk "
                "Microsoft 365) telah terpasang dan berlisensi aktif pada komputer ini."
            ) from exc

        self._configure_application()
        LOGGER.info("Sesi Microsoft Word dibuka.")
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def close(self) -> None:
        if self._app is not None:
            try:
                self._app.Quit(WD_DO_NOT_SAVE_CHANGES)
            except Exception as exc:  # noqa: BLE001
                LOGGER.warning("Gagal menutup Word secara normal: %s", _com_message(exc))
            finally:
                self._app = None
        self._safe_couninitialize()
        LOGGER.info("Sesi Microsoft Word ditutup.")

    def _safe_couninitialize(self) -> None:
        if self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:  # noqa: BLE001
                pass

    def _configure_application(self) -> None:
        """Silence every interactive behaviour that could block automation."""
        app = self._app
        for attribute, value in (
            ("Visible", False),
            ("DisplayAlerts", 0),
            ("ScreenUpdating", False),
            ("AutomationSecurity", MSO_AUTOMATION_SECURITY_FORCE_DISABLE),
        ):
            try:
                setattr(app, attribute, value)
            except Exception as exc:  # noqa: BLE001
                LOGGER.debug("Properti %s tidak dapat diatur: %s", attribute, _com_message(exc))

        for option, value in (
            ("Pagination", True),          # background repagination must be on
            ("UpdateLinksAtOpen", False),
            ("SavePropertiesPrompt", False),
            ("ConfirmConversions", False),
            ("WarnBeforeSavingPrintingSendingMarkup", False),
        ):
            try:
                setattr(app.Options, option, value)
            except Exception as exc:  # noqa: BLE001
                LOGGER.debug("Opsi %s tidak dapat diatur: %s", option, _com_message(exc))

    # -- document handling -------------------------------------------------

    @contextmanager
    def open_document(self, path: Path, read_only: bool) -> Iterator[Any]:
        """Open a document and guarantee that it is closed again."""
        if self._app is None:
            raise WordEngineError("Sesi Microsoft Word belum aktif.")
        document = None
        try:
            document = self._app.Documents.Open(
                FileName=str(path),
                ConfirmConversions=False,
                ReadOnly=read_only,
                AddToRecentFiles=False,
                Revert=True,
                Visible=False,
                OpenAndRepair=False,
                NoEncodingDialog=True,
            )
        except Exception as exc:  # noqa: BLE001
            message = _com_message(exc)
            LOGGER.error("Word gagal membuka %s: %s", path, message)
            raise DocumentOpenError(
                "Microsoft Word tidak dapat membuka dokumen ini. "
                "Dokumen mungkin rusak, terproteksi kata sandi, atau sedang "
                f"dikunci oleh proses lain.\n\nDetail: {message}"
            ) from exc

        try:
            yield document
        finally:
            try:
                document.Close(WD_DO_NOT_SAVE_CHANGES)
            except Exception as exc:  # noqa: BLE001
                LOGGER.warning("Dokumen gagal ditutup: %s", _com_message(exc))

    # -- pagination --------------------------------------------------------

    @staticmethod
    def repaginate(document: Any) -> None:
        try:
            document.Repaginate()
        except Exception as exc:  # noqa: BLE001
            LOGGER.warning("Repaginate gagal: %s", _com_message(exc))

    def page_count(self, document: Any) -> int:
        """Return the rendered page count reported by Word."""
        self.repaginate(document)
        try:
            pages = int(document.ComputeStatistics(WD_STATISTIC_PAGES))
        except Exception as exc:  # noqa: BLE001
            raise PaginationError(
                "Microsoft Word tidak dapat menghitung jumlah halaman dokumen ini.\n\n"
                f"Detail: {_com_message(exc)}"
            ) from exc
        if pages < 1:
            raise PaginationError(
                "Jumlah halaman yang dilaporkan Microsoft Word tidak masuk akal. "
                "Dokumen kemungkinan kosong atau rusak."
            )
        return pages

    def page_start_offset(self, document: Any, page: int) -> int:
        """Return the character offset at which the given page begins.

        The offset is obtained from Word's own layout engine through GoTo, which
        is the authoritative mapping between rendered pages and the character
        stream of the document.
        """
        try:
            target = document.GoTo(
                What=WD_GOTO_PAGE, Which=WD_GOTO_ABSOLUTE, Count=page
            )
            offset = int(target.Start)
        except Exception as exc:  # noqa: BLE001
            raise PaginationError(
                f"Posisi awal halaman {page} tidak dapat ditentukan oleh Microsoft Word.\n\n"
                f"Detail: {_com_message(exc)}"
            ) from exc

        try:
            content_end = int(document.Content.End)
            content_start = int(document.Content.Start)
        except Exception as exc:  # noqa: BLE001
            raise PaginationError(
                f"Rentang isi dokumen tidak dapat dibaca.\n\nDetail: {_com_message(exc)}"
            ) from exc

        if not content_start < offset < content_end:
            raise PaginationError(
                f"Titik pisah pada halaman {page} berada di batas dokumen sehingga "
                "salah satu bagian akan kosong. Pilih kombinasi halaman lain."
            )
        return offset

    # -- editing -----------------------------------------------------------

    @staticmethod
    def delete_range(document: Any, start: int, end: int) -> None:
        """Delete a character range while leaving the remaining content intact."""
        if end <= start:
            return
        try:
            document.Range(Start=start, End=end).Delete()
        except Exception as exc:  # noqa: BLE001
            raise WordEngineError(
                "Bagian dokumen gagal dihapus saat membentuk hasil pemisahan.\n\n"
                f"Detail: {_com_message(exc)}"
            ) from exc

    @staticmethod
    def save_as_docx(document: Any, target: Path) -> None:
        """Persist the document as a .docx package."""
        try:
            document.SaveAs2(
                FileName=str(target),
                FileFormat=WD_FORMAT_DOCUMENT_DEFAULT,
                AddToRecentFiles=False,
            )
        except AttributeError:
            try:
                document.SaveAs(
                    FileName=str(target),
                    FileFormat=WD_FORMAT_DOCUMENT_DEFAULT,
                    AddToRecentFiles=False,
                )
            except Exception as exc:  # noqa: BLE001
                raise SaveError(_save_message(target, exc)) from exc
        except Exception as exc:  # noqa: BLE001
            raise SaveError(_save_message(target, exc)) from exc

    # -- verification ------------------------------------------------------

    def verify_document(self, path: Path) -> int:
        """Reopen a produced file with Word and return its page count.

        A file that Word itself refuses to reopen is not a successful output,
        regardless of what SaveAs reported.
        """
        with self.open_document(path, read_only=True) as document:
            return self.page_count(document)


def _save_message(target: Path, exc: BaseException) -> str:
    return (
        f"Microsoft Word gagal menyimpan file '{target.name}'. "
        "Penyebab yang umum adalah izin folder yang tidak mencukupi, ruang disk "
        "yang habis, atau file tujuan sedang terbuka di aplikasi lain.\n\n"
        f"Detail: {_com_message(exc)}"
    )


def word_is_available() -> tuple[bool, str]:
    """Probe whether Microsoft Word can be started on this machine."""
    try:
        with WordSession():
            return True, ""
    except WordEngineError as exc:
        return False, str(exc)
    except Exception as exc:  # noqa: BLE001
        return False, f"Pemeriksaan Microsoft Word gagal: {exc}"

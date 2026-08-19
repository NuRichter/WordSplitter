"""Orchestration of the split operation.

The engine follows a staged commit protocol. Every artefact is produced under a
temporary name on the same volume as the destination, verified, and only then
promoted to its final name with an atomic replace. The original document is
never touched until both outputs exist and have been reopened successfully by
Microsoft Word.
"""

from __future__ import annotations

import logging
import os
import shutil
import threading
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, List, Optional

import utils
import validation
from validation import SplitPlan, ValidationError
from word_engine import (
    PaginationError,
    WordEngineError,
    WordNotAvailableError,
    WordSession,
)

LOGGER = logging.getLogger("WordSplitter.split")

ProgressCallback = Callable[[str], None]


class OperationCancelled(Exception):
    """Raised when the user cancelled the operation between two stages."""


class SplitError(Exception):
    """Any recoverable failure that must be reported to the user verbatim."""


@dataclass
class SplitRequest:
    source: Path
    page_a: int
    page_b: int
    output_one: Path
    output_two: Path
    delete_original: bool = False
    overwrite_allowed: bool = False


@dataclass
class SplitResult:
    plan: SplitPlan
    output_one: Path
    output_two: Path
    pages_one: int
    pages_two: int
    original_deleted: bool
    warnings: List[str] = field(default_factory=list)


class AnalysisResult:
    """Outcome of the pagination analysis stage."""

    def __init__(self, total_pages: int) -> None:
        self.total_pages = total_pages


def _check_cancel(cancel_event: Optional[threading.Event]) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise OperationCancelled()


def _report(progress: Optional[ProgressCallback], message: str) -> None:
    LOGGER.info(message)
    if progress is not None:
        progress(message)


def analyze_document(
    source: Path,
    progress: Optional[ProgressCallback] = None,
    cancel_event: Optional[threading.Event] = None,
) -> AnalysisResult:
    """Open the document with Word and report its rendered page count."""
    _report(progress, "Memeriksa integritas dokumen...")
    valid, reason = utils.is_valid_docx_container(source)
    if not valid:
        raise SplitError(reason)

    _check_cancel(cancel_event)
    _report(progress, "Menjalankan Microsoft Word untuk menganalisis pagination...")

    with WordSession() as session:
        _check_cancel(cancel_event)
        with session.open_document(source, read_only=True) as document:
            pages = session.page_count(document)

    _report(progress, f"Analisis selesai. Dokumen memiliki {pages} halaman.")
    return AnalysisResult(total_pages=pages)


def _preflight(request: SplitRequest) -> None:
    """Checks that must pass before Word is started at all."""
    if not request.source.exists():
        raise SplitError(
            "File asli tidak ditemukan lagi. Kemungkinan telah dipindahkan, "
            "diubah namanya, atau dihapus setelah dipilih."
        )

    readable, reason = utils.file_is_readable(request.source)
    if not readable:
        raise SplitError(reason)

    valid, reason = utils.is_valid_docx_container(request.source)
    if not valid:
        raise SplitError(reason)

    writable, reason = utils.directory_is_writable(request.source.parent)
    if not writable:
        raise SplitError(reason)

    source_size = request.source.stat().st_size
    required = validation.estimate_required_space(source_size)
    free = utils.free_space_bytes(request.source.parent)
    if free is not None and free < required:
        raise SplitError(
            "Ruang disk tidak mencukupi untuk menyelesaikan operasi dengan aman. "
            f"Dibutuhkan sekitar {utils.human_readable_size(required)}, "
            f"tersedia {utils.human_readable_size(free)}."
        )

    if not request.overwrite_allowed:
        conflicts = validation.existing_outputs(request.output_one, request.output_two)
        if conflicts:
            raise SplitError(
                "File output sudah ada: "
                + ", ".join(path.name for path in conflicts)
            )


def _build_part(
    session: WordSession,
    source: Path,
    temp_path: Path,
    delete_start: int,
    delete_end: int,
    label: str,
    progress: Optional[ProgressCallback],
) -> None:
    """Copy the source and remove the range that does not belong to this part."""
    _report(progress, f"Menyiapkan {label}...")
    try:
        shutil.copy2(source, temp_path)
    except OSError as exc:
        raise SplitError(
            f"Salinan kerja untuk {label} gagal dibuat: {exc.strerror or exc}"
        ) from exc

    with session.open_document(temp_path, read_only=False) as document:
        session.repaginate(document)
        session.delete_range(document, delete_start, delete_end)
        session.save_as_docx(document, temp_path)
    _report(progress, f"{label} selesai dibentuk.")


def _verify_part(session: WordSession, path: Path, label: str) -> int:
    """Structural and semantic verification of a produced file."""
    if not path.exists():
        raise SplitError(f"{label} tidak ditemukan setelah proses penyimpanan.")
    if path.stat().st_size == 0:
        raise SplitError(f"{label} berukuran 0 byte sehingga dianggap gagal.")

    valid, reason = utils.is_valid_docx_container(path)
    if not valid:
        raise SplitError(f"{label} bukan paket .docx yang valid. {reason}")

    try:
        pages = session.verify_document(path)
    except WordEngineError as exc:
        raise SplitError(
            f"{label} tidak dapat dibuka kembali oleh Microsoft Word untuk validasi.\n\n{exc}"
        ) from exc
    if pages < 1:
        raise SplitError(f"{label} tidak memiliki halaman yang dapat dirender.")
    return pages


def _commit(temp_path: Path, final_path: Path, label: str) -> None:
    """Promote a verified temporary file to its final name atomically."""
    try:
        os.replace(temp_path, final_path)
    except PermissionError as exc:
        raise SplitError(
            f"{label} tidak dapat disimpan sebagai '{final_path.name}'. "
            "File tujuan sedang terbuka di aplikasi lain atau izin folder tidak mencukupi."
        ) from exc
    except OSError as exc:
        raise SplitError(
            f"{label} gagal dipindahkan ke nama final: {exc.strerror or exc}"
        ) from exc


def perform_split(
    request: SplitRequest,
    progress: Optional[ProgressCallback] = None,
    cancel_event: Optional[threading.Event] = None,
) -> SplitResult:
    """Execute the full split protocol and return a verified result."""
    warnings: List[str] = []
    temp_one: Optional[Path] = None
    temp_two: Optional[Path] = None
    committed_one = False
    committed_two = False

    _report(progress, "Memeriksa prasyarat...")
    _preflight(request)
    _check_cancel(cancel_event)

    # A pre existing file at the Part 1 name belongs to the user. If the run
    # fails after Part 1 has been committed, that file is left alone rather than
    # deleted, because removing it would destroy data the user already had.
    part_one_preexisted = request.output_one.exists()

    try:
        with WordSession() as session:
            _check_cancel(cancel_event)

            _report(progress, "Menganalisis pagination dokumen...")
            with session.open_document(request.source, read_only=True) as document:
                total_pages = session.page_count(document)
                plan = validation.compute_split_plan(
                    request.page_a, request.page_b, total_pages
                )
                split_offset = session.page_start_offset(document, plan.split_page)
                content_start = int(document.Content.Start)
                content_end = int(document.Content.End)

            LOGGER.info(
                "Rencana pisah: %s, offset karakter %d, rentang isi %d-%d",
                plan.description,
                split_offset,
                content_start,
                content_end,
            )
            _report(progress, f"Titik pisah ditentukan. {plan.description}.")
            _check_cancel(cancel_event)

            temp_one = utils.temp_sibling_path(request.output_one, "part1")
            temp_two = utils.temp_sibling_path(request.output_two, "part2")

            # Part 1 keeps the content before the split point.
            _build_part(
                session, request.source, temp_one,
                split_offset, content_end, "Part 1", progress,
            )
            _check_cancel(cancel_event)

            # Part 2 keeps the content from the split point onwards.
            _build_part(
                session, request.source, temp_two,
                content_start, split_offset, "Part 2", progress,
            )
            _check_cancel(cancel_event)

            _report(progress, "Memvalidasi hasil...")
            pages_one = _verify_part(session, temp_one, "Part 1")
            pages_two = _verify_part(session, temp_two, "Part 2")

            # A silent failure of the delete step would produce two copies of
            # the whole document. Both parts reaching the original page count is
            # the signature of that failure.
            if pages_one >= total_pages and pages_two >= total_pages:
                raise SplitError(
                    "Kedua file hasil masih memuat seluruh isi dokumen, sehingga "
                    "pemisahan dianggap gagal. Proses dibatalkan dan tidak ada "
                    "file yang ditulis."
                )
            if utils.files_are_identical(temp_one, temp_two):
                warnings.append(
                    "Kedua file hasil memiliki isi biner yang identik. "
                    "Periksa hasilnya sebelum digunakan."
                )

        # Word is closed before the commit so that no handle remains open.
        _check_cancel(cancel_event)
        _report(progress, "Menyimpan file hasil...")
        _commit(temp_one, request.output_one, "Part 1")
        committed_one = True
        _commit(temp_two, request.output_two, "Part 2")
        committed_two = True

    except OperationCancelled:
        utils.cleanup_temp_artifact(temp_one if not committed_one else None)
        utils.cleanup_temp_artifact(temp_two if not committed_two else None)
        raise
    except (SplitError, ValidationError):
        utils.cleanup_temp_artifact(temp_one if not committed_one else None)
        utils.cleanup_temp_artifact(temp_two if not committed_two else None)
        if committed_one and not committed_two and not part_one_preexisted:
            utils.safe_unlink(request.output_one)
        raise
    except WordNotAvailableError as exc:
        utils.cleanup_temp_artifact(temp_one)
        utils.cleanup_temp_artifact(temp_two)
        raise SplitError(str(exc)) from exc
    except (PaginationError, WordEngineError) as exc:
        utils.cleanup_temp_artifact(temp_one if not committed_one else None)
        utils.cleanup_temp_artifact(temp_two if not committed_two else None)
        if committed_one and not committed_two and not part_one_preexisted:
            utils.safe_unlink(request.output_one)
        raise SplitError(str(exc)) from exc
    except Exception as exc:  # noqa: BLE001 - converted into a user level error
        LOGGER.exception("Kegagalan tak terduga selama proses pemisahan.")
        utils.cleanup_temp_artifact(temp_one if not committed_one else None)
        utils.cleanup_temp_artifact(temp_two if not committed_two else None)
        if committed_one and not committed_two and not part_one_preexisted:
            utils.safe_unlink(request.output_one)
        raise SplitError(
            "Terjadi kesalahan tak terduga selama proses pemisahan. "
            "Detail teknis telah dicatat pada file log."
        ) from exc

    original_deleted = False
    if request.delete_original:
        _report(progress, "Menghapus file asli...")
        try:
            request.source.unlink()
            original_deleted = True
        except PermissionError:
            warnings.append(
                "File asli tidak dapat dihapus karena sedang digunakan oleh aplikasi "
                "lain atau izin tidak mencukupi. Kedua file hasil tetap tersimpan "
                "dengan benar."
            )
        except OSError as exc:
            warnings.append(
                "File asli tidak dapat dihapus: "
                f"{exc.strerror or exc}. Kedua file hasil tetap tersimpan dengan benar."
            )

    result = SplitResult(
        plan=plan,
        output_one=request.output_one,
        output_two=request.output_two,
        pages_one=pages_one,
        pages_two=pages_two,
        original_deleted=original_deleted,
        warnings=warnings,
    )
    LOGGER.info(
        "Pemisahan berhasil. %s | Part 1: %d halaman | Part 2: %d halaman | original dihapus: %s",
        plan.description, pages_one, pages_two, original_deleted,
    )
    return result

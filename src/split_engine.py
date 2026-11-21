"""Orchestration of analysis and split (runs inside the worker process).

Pipeline
--------
1. The source is checked without Word (readable, sound .docx package) and its
   fingerprint is compared with the one the user saw during analysis.
2. A byte copy of the source ("snapshot") is placed in a private job folder
   under a name Word has never seen. Word only ever opens that snapshot, so
   the user's file is never locked, never shown in Word's recent list and
   never touched by Word's crash recovery, even when the user has it open.
3. One Word session does everything: count pages, find the cut, build Part 1
   on the already open snapshot, build Part 2 on a fresh open, check that
   Word reopens both files.
4. Results are committed transactionally: files already at the target names
   are moved aside first, both parts are copied into place, and on any
   failure the previous state is restored. The original is removed last, and
   only into the Recycle Bin when the drive has one.
"""

from __future__ import annotations

import logging
import os
import shutil
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

import utils
import validation
import winsys
from errors import AppError, OperationCancelled
from validation import SplitPlan

LOGGER = logging.getLogger("WordSplitter.split")

Emit = Callable[..., None]

# Set while results are being copied into place. A cancel that arrives now is
# deliberately ignored: stopping half way would be worse than finishing.
COMMITTING = threading.Event()


@dataclass
class SplitRequest:
    source: Path
    job_dir: Path
    fingerprint: Optional[utils.Fingerprint]
    total_pages: int
    page_a: int
    page_b: int
    output_one: Path
    output_two: Path
    delete_original: bool = False
    overwrite_allowed: bool = False

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "SplitRequest":
        fingerprint = data.get("fingerprint")
        return cls(
            source=Path(data["source"]),
            job_dir=Path(data["job_dir"]),
            fingerprint=utils.Fingerprint.from_list(fingerprint) if fingerprint else None,
            total_pages=int(data.get("total_pages") or 0),
            page_a=int(data["page_a"]),
            page_b=int(data["page_b"]),
            output_one=Path(data["output_one"]),
            output_two=Path(data["output_two"]),
            delete_original=bool(data.get("delete_original")),
            overwrite_allowed=bool(data.get("overwrite_allowed")),
        )


@dataclass
class SplitResult:
    plan: SplitPlan
    output_one: Path
    output_two: Path
    pages_one: int
    pages_two: int
    original_removed: str = "kept"  # kept | recycled | deleted | failed
    warnings: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "plan": [self.plan.page_a, self.plan.page_b, self.plan.split_page, self.plan.total_pages],
            "output_one": str(self.output_one),
            "output_two": str(self.output_two),
            "pages_one": self.pages_one,
            "pages_two": self.pages_two,
            "original_removed": self.original_removed,
            "warnings": self.warnings,
        }


def _progress(emit: Optional[Emit], key: str, pct: int, cancel: Optional[threading.Event] = None) -> None:
    if cancel is not None and cancel.is_set():
        raise OperationCancelled()
    LOGGER.info("Stage %s (%d%%)", key, pct)
    if emit is not None:
        emit("progress", key=key, pct=pct)


def _session_factory():
    import word_engine

    return word_engine.WordSession


def _snapshot(source: Path, job_dir: Path, expected: Optional[utils.Fingerprint]) -> Path:
    """Copy the source into the job folder and prove the copy is that version."""
    for _attempt in range(3):
        before = utils.Fingerprint.of(source)
        if expected is not None and before != expected:
            raise AppError("err.source_changed")
        target = job_dir / utils.unique_name("source")
        try:
            shutil.copyfile(source, target)
        except PermissionError:
            raise AppError("err.in_use") from None
        except OSError as exc:
            raise AppError("err.read_failed", detail=str(exc.strerror or exc)) from None
        if utils.Fingerprint.of(source) == before and target.stat().st_size == before.size:
            return target
        utils.safe_unlink(target)  # the file was being saved while we copied it
    raise AppError("err.source_changed")


def analyze(source: Path, job_dir: Path, emit: Optional[Emit] = None,
            job: Optional[winsys.KillOnCloseJob] = None, session_factory=None,
            cancel: Optional[threading.Event] = None) -> Dict[str, Any]:
    """Count the rendered pages of a document."""
    _progress(emit, "progress.checking", 5, cancel)
    utils.check_readable(source)
    utils.check_docx_container(source)
    fingerprint = utils.Fingerprint.of(source)

    _progress(emit, "progress.copying", 10, cancel)
    job_dir.mkdir(parents=True, exist_ok=True)
    snapshot = _snapshot(source, job_dir, fingerprint)

    factory = session_factory or _session_factory()
    _progress(emit, "progress.starting_word", 20, cancel)
    session = factory(job)
    with session:
        _progress(emit, "progress.opening", 45, cancel)
        document = session.open(snapshot)
        try:
            _progress(emit, "progress.paginating", 70, cancel)
            pages = session.page_count(document)
        finally:
            session.close_document(document)
    LOGGER.info("Analysis: %s has %d pages.", source.name, pages)
    return {"pages": pages, "fingerprint": fingerprint.as_list()}


def _preflight(request: SplitRequest) -> None:
    source = request.source
    if not source.exists():
        raise AppError("err.not_found")
    utils.check_readable(source)
    utils.check_docx_container(source)
    if request.fingerprint is not None and utils.Fingerprint.of(source) != request.fingerprint:
        raise AppError("err.source_changed")

    destination = request.output_one.parent
    utils.check_writable_directory(destination)
    size = source.stat().st_size
    needed = validation.estimate_required_space(size)
    for folder in {destination, request.job_dir.parent}:
        free = utils.free_space_bytes(folder)
        if free is not None and free < needed:
            raise AppError("err.disk_space", needed_bytes=needed, free_bytes=free)

    if not request.overwrite_allowed:
        conflicts = validation.existing_outputs(request.output_one, request.output_two)
        if conflicts:
            raise AppError("err.outputs_exist", names=", ".join(p.name for p in conflicts))


def _expect(warnings: List[Dict[str, Any]], label: str, actual: int, expected: int) -> None:
    if actual != expected:
        warnings.append({"key": "warn.pages_differ", "params": {"label": label, "actual": actual, "expected": expected}})


def _commit(parts: List[tuple], overwrite: bool, emit: Optional[Emit]) -> None:
    """Copy both verified parts into place; all or nothing."""
    COMMITTING.set()
    if emit is not None:
        emit("commit")  # from here on the GUI no longer offers Cancel
        time.sleep(0.15)  # let the GUI register that before any file moves
    backups: List[tuple] = []
    written: List[Path] = []
    try:
        for _source, final, _label in parts:
            if final.exists():
                if not overwrite:
                    raise AppError("err.outputs_exist", names=final.name)
                backup = utils.temp_sibling_path(final)
                os.replace(final, backup)
                backups.append((backup, final))
        for source, final, label in parts:
            try:
                utils.copy_into_place(source, final, overwrite=False)
            except PermissionError as exc:
                raise AppError("err.commit_failed", detail=str(exc), label=label, name=final.name) from exc
            except OSError as exc:
                raise AppError("err.commit_failed", detail=str(exc.strerror or exc), label=label, name=final.name) from exc
            written.append(final)
    except BaseException:
        for final in written:
            utils.safe_unlink(final)
        for backup, final in backups:
            try:
                os.replace(backup, final)
            except OSError:
                LOGGER.exception("Could not restore %s from %s", final, backup)
        raise
    for backup, _final in backups:
        utils.safe_unlink(backup)


def _remove_original(source: Path) -> str:
    if winsys.drive_has_recycle_bin(str(source)):
        if winsys.move_to_recycle_bin(str(source)):
            return "recycled"
        return "failed"
    try:
        source.unlink()
        return "deleted"
    except OSError:
        return "failed"


def split(request: SplitRequest, emit: Optional[Emit] = None,
          job: Optional[winsys.KillOnCloseJob] = None, session_factory=None,
          cancel: Optional[threading.Event] = None) -> SplitResult:
    """Execute the full split protocol and return a verified result."""
    warnings: List[Dict[str, Any]] = []
    _progress(emit, "progress.checking", 3, cancel)
    _preflight(request)

    _progress(emit, "progress.copying", 6, cancel)
    request.job_dir.mkdir(parents=True, exist_ok=True)
    snapshot = _snapshot(request.source, request.job_dir, request.fingerprint)
    temp_one = request.job_dir / utils.unique_name("part1")
    temp_two = request.job_dir / utils.unique_name("part2")

    factory = session_factory or _session_factory()
    _progress(emit, "progress.starting_word", 12, cancel)
    session = factory(job)
    with session:
        _progress(emit, "progress.paginating", 22, cancel)
        document = session.open(snapshot)
        try:
            total = session.page_count(document)
            if request.total_pages and total != request.total_pages:
                raise AppError("err.source_changed", detail=f"pages {request.total_pages} -> {total}")
            plan = validation.compute_split_plan(request.page_a, request.page_b, total)
            split_offset = session.page_start(document, plan.split_page)
            LOGGER.info("Plan: %s, cut at character %d", plan.description, split_offset)

            _progress(emit, "progress.part1", 40, cancel)
            report_one = session.cut_keep_before(document, split_offset)
            pages_one = session.page_count(document)
            session.save_docx(document, temp_one)
        finally:
            session.close_document(document)

        _progress(emit, "progress.part2", 62, cancel)
        document = session.open(snapshot)
        try:
            report_two = session.cut_keep_from(document, split_offset)
            pages_two = session.page_count(document)
            session.save_docx(document, temp_two)
        finally:
            session.close_document(document)

        _progress(emit, "progress.verifying", 82, cancel)
        session.reopen_check(temp_one, "result.file1")
        session.reopen_check(temp_two, "result.file2")

    LOGGER.info("Part 1 steps: %s", getattr(report_one, "steps", []))
    LOGGER.info("Part 2 steps: %s", getattr(report_two, "steps", []))
    for path, label in ((temp_one, "result.file1"), (temp_two, "result.file2")):
        try:
            utils.check_docx_container(path)
        except AppError as exc:
            raise AppError("err.verify_failed", detail=exc.key, label=label) from None

    if pages_one >= total and pages_two >= total:
        raise AppError("err.both_full")
    _expect(warnings, "result.file1", pages_one, plan.split_page - 1)
    _expect(warnings, "result.file2", pages_two, total - plan.split_page + 1)
    if getattr(report_one, "fields_unlinked", 0) or getattr(report_two, "fields_unlinked", 0):
        warnings.append({"key": "warn.fields_unlinked", "params": {}})
    if utils.files_are_identical(temp_one, temp_two):
        warnings.append({"key": "warn.identical", "params": {}})

    _progress(emit, "progress.saving", 92, cancel)
    _commit(
        [(temp_one, request.output_one, "result.file1"), (temp_two, request.output_two, "result.file2")],
        request.overwrite_allowed,
        emit,
    )

    original = "kept"
    if request.delete_original:
        _progress(emit, "progress.removing_original", 97)
        original = _remove_original(request.source)
        if original == "failed":
            warnings.append({"key": "warn.delete_failed", "params": {}})

    result = SplitResult(
        plan=plan,
        output_one=request.output_one,
        output_two=request.output_two,
        pages_one=pages_one,
        pages_two=pages_two,
        original_removed=original,
        warnings=warnings,
    )
    LOGGER.info("Split done. %s | part1=%d part2=%d | original=%s | warnings=%s",
                plan.description, pages_one, pages_two, original, warnings)
    return result

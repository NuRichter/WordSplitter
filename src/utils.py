"""General purpose helpers shared by the GUI and the worker.

Nothing in this module depends on Microsoft Word or on the GUI toolkit, which
keeps it unit testable on any platform.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import time
import uuid
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Tuple

from errors import AppError

DOCX_EXTENSION = ".docx"
REQUIRED_ZIP_ENTRIES = ("[Content_Types].xml", "word/document.xml")
# Encrypted Office documents and legacy .doc files are OLE compound files.
CFB_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
ZIP_SIGNATURE = b"PK"
# Word cannot open paths longer than this, even on systems with long path support.
WORD_MAX_PATH = 255


@dataclass(frozen=True)
class FileInfo:
    """Immutable description of a file on disk."""

    path: Path
    size_bytes: int

    @property
    def name(self) -> str:
        return self.path.name

    @property
    def directory(self) -> Path:
        return self.path.parent


@dataclass(frozen=True)
class Fingerprint:
    """Cheap identity of a file version: size plus modification time."""

    size: int
    mtime_ns: int

    @classmethod
    def of(cls, path: Path) -> "Fingerprint":
        stat = path.stat()
        return cls(size=stat.st_size, mtime_ns=stat.st_mtime_ns)

    def as_list(self) -> list:
        return [self.size, self.mtime_ns]

    @classmethod
    def from_list(cls, values) -> "Fingerprint":
        return cls(size=int(values[0]), mtime_ns=int(values[1]))


def format_size(size: int, units: Tuple[str, str, str, str, str] = ("B", "KB", "MB", "GB", "TB"),
                decimal: str = ".") -> str:
    """Render a byte count with binary multiples and a locale decimal mark."""
    if size < 0:
        return "?"
    value = float(size)
    index = 0
    while value >= 1024.0 and index < len(units) - 1:
        value /= 1024.0
        index += 1
    if index == 0:
        return f"{int(value)} {units[0]}"
    return f"{value:.2f}".replace(".", decimal) + f" {units[index]}"


def check_docx_container(path: Path) -> None:
    """Verify that a file is a structurally sound Office Open XML package.

    This runs before Word is involved at all. It catches truncated downloads,
    renamed .doc binaries, password protected packages and damaged archives,
    each with its own explanation, and raises AppError on the first problem.
    """
    try:
        if not path.is_file():
            raise AppError("err.not_found")
        size = path.stat().st_size
        if size == 0:
            raise AppError("err.empty_file")
        with path.open("rb") as handle:
            head = handle.read(8)
        if head.startswith(CFB_SIGNATURE):
            # Either an old .doc renamed to .docx or an encrypted .docx package.
            raise AppError("err.encrypted_or_legacy")
        if not head.startswith(ZIP_SIGNATURE):
            raise AppError("err.not_zip")
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            for required in REQUIRED_ZIP_ENTRIES:
                if required not in names:
                    raise AppError("err.zip_missing", entry=required)
            # CRC check of the entries Word reads first; a full testzip of a
            # large image heavy document would cost seconds for little gain.
            for required in REQUIRED_ZIP_ENTRIES:
                with archive.open(required) as entry:
                    while entry.read(1 << 16):
                        pass
    except AppError:
        raise
    except (zipfile.BadZipFile, zipfile.LargeZipFile, EOFError):
        raise AppError("err.zip_corrupt") from None
    except PermissionError:
        raise AppError("err.in_use") from None
    except OSError as exc:
        raise AppError("err.read_failed", detail=str(exc.strerror or exc)) from None
    except Exception as exc:  # noqa: BLE001 - e.g. zlib errors from a damaged entry
        raise AppError("err.zip_corrupt", detail=type(exc).__name__) from None


def check_readable(path: Path) -> None:
    """Check that the file can be opened for reading right now."""
    try:
        with path.open("rb") as handle:
            handle.read(1)
    except PermissionError:
        raise AppError("err.in_use") from None
    except FileNotFoundError:
        raise AppError("err.not_found") from None
    except OSError as exc:
        raise AppError("err.read_failed", detail=str(exc.strerror or exc)) from None


def check_writable_directory(directory: Path) -> None:
    """Verify write permission by creating and removing a probe file."""
    if not directory.is_dir():
        raise AppError("err.dest_missing")
    probe = directory / f".wordsplitter-probe-{uuid.uuid4().hex[:8]}.tmp"
    try:
        with probe.open("wb") as handle:
            handle.write(b"0")
    except PermissionError:
        raise AppError("err.dest_readonly") from None
    except OSError as exc:
        raise AppError("err.dest_readonly", detail=str(exc.strerror or exc)) from None
    finally:
        safe_unlink(probe)


def free_space_bytes(directory: Path) -> Optional[int]:
    try:
        return shutil.disk_usage(directory).free
    except OSError:
        return None


def ensure_docx_extension(name: str) -> str:
    """Append .docx when the user omitted it, normalising its case."""
    cleaned = name.strip()
    if cleaned.lower().endswith(DOCX_EXTENSION):
        return cleaned[: -len(DOCX_EXTENSION)] + DOCX_EXTENSION
    return cleaned + DOCX_EXTENSION


def strip_docx_extension(name: str) -> str:
    cleaned = name.strip()
    if cleaned.lower().endswith(DOCX_EXTENSION):
        return cleaned[: -len(DOCX_EXTENSION)]
    return cleaned


def jobs_root() -> Path:
    """Private scratch area for working copies. Short, local and ASCII only."""
    return Path(tempfile.gettempdir()) / "WordSplitter"


def new_job_dir() -> Path:
    """Create a fresh, uniquely named job directory."""
    root = jobs_root()
    root.mkdir(parents=True, exist_ok=True)
    path = root / f"job-{os.getpid()}-{uuid.uuid4().hex[:10]}"
    path.mkdir()
    return path


def unique_name(prefix: str, suffix: str = ".docx") -> str:
    """A file name Word has never seen before.

    Word remembers documents that were open when it was terminated and asks
    about them the next time a file with that name is opened. That question
    blocks automation, so every working file gets a name that is never reused.
    """
    return f"{prefix}-{uuid.uuid4().hex[:12]}{suffix}"


def temp_sibling_path(target: Path) -> Path:
    """A unique hidden temporary name next to the final target (same volume)."""
    return target.parent / f"~wordsplitter-{uuid.uuid4().hex[:12]}.tmp"


def safe_unlink(path: Optional[Path]) -> None:
    """Delete a file, ignoring every failure. Used only for temporary artefacts."""
    if path is None:
        return
    try:
        path.unlink(missing_ok=True)
    except Exception:  # noqa: BLE001 - cleanup must never raise
        pass


def word_owner_files(path: Path) -> list:
    """Return the possible Word lock file names ("~$...") for a document."""
    name = path.name
    candidates = {f"~${name}"}
    if len(name) > 2:
        candidates.add(f"~${name[2:]}")
    return [path.parent / candidate for candidate in candidates]


def remove_tree(path: Optional[Path], attempts: int = 20, delay: float = 0.15) -> bool:
    """Remove a job directory, retrying while a terminated Word releases handles."""
    if path is None:
        return True
    for _ in range(max(1, attempts)):
        try:
            shutil.rmtree(path)
            return True
        except FileNotFoundError:
            return True
        except OSError:
            time.sleep(delay)
    return not path.exists()


def remove_files(paths: Iterable[Optional[Path]], attempts: int = 20, delay: float = 0.15) -> None:
    pending = [p for p in paths if p is not None]
    for _ in range(max(1, attempts)):
        remaining = []
        for item in pending:
            try:
                item.unlink(missing_ok=True)
            except OSError:
                remaining.append(item)
        if not remaining:
            return
        pending = remaining
        time.sleep(delay)


def copy_into_place(source: Path, final: Path, overwrite: bool) -> None:
    """Copy a verified file into its final location as atomically as possible.

    The bytes are written to a temporary sibling first, flushed to disk, and
    only then renamed onto the final name. Copying (instead of moving) lets
    the new file inherit the permissions of the destination folder.
    """
    staging = temp_sibling_path(final)
    try:
        with source.open("rb") as reader, staging.open("wb") as writer:
            shutil.copyfileobj(reader, writer, 1 << 20)
            writer.flush()
            os.fsync(writer.fileno())
        if not overwrite and final.exists():
            raise FileExistsError(str(final))
        os.replace(staging, final)
    except BaseException:
        safe_unlink(staging)
        raise


def files_are_identical(first: Path, second: Path) -> bool:
    """Byte comparison used to detect a split that produced two equal outputs."""
    try:
        if first.stat().st_size != second.stat().st_size:
            return False
        with first.open("rb") as fa, second.open("rb") as fb:
            while True:
                chunk_a = fa.read(1 << 16)
                chunk_b = fb.read(1 << 16)
                if chunk_a != chunk_b:
                    return False
                if not chunk_a:
                    return True
    except OSError:
        return False

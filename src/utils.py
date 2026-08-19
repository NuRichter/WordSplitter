"""General purpose helpers shared across the application.

Nothing in this module depends on Microsoft Word or on the GUI toolkit, which
keeps it unit testable on any platform.
"""

from __future__ import annotations

import os
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

DOCX_EXTENSION = ".docx"
REQUIRED_ZIP_ENTRIES = ("[Content_Types].xml", "word/document.xml")


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

    @property
    def size_human(self) -> str:
        return human_readable_size(self.size_bytes)


def human_readable_size(size: int) -> str:
    """Render a byte count using binary multiples."""
    if size < 0:
        return "unknown"
    units = ("B", "KB", "MB", "GB", "TB")
    value = float(size)
    index = 0
    while value >= 1024.0 and index < len(units) - 1:
        value /= 1024.0
        index += 1
    if index == 0:
        return f"{int(value)} {units[index]}"
    return f"{value:.2f} {units[index]}"


def is_valid_docx_container(path: Path) -> Tuple[bool, str]:
    """Verify that a file is a structurally sound Office Open XML package.

    This is a cheap structural check performed without Microsoft Word. It
    detects truncated downloads, renamed .doc binaries and corrupted archives
    before any COM call is attempted.
    """
    try:
        if not path.is_file():
            return False, "File tidak ditemukan."
        if path.stat().st_size == 0:
            return False, "File berukuran 0 byte."
        with zipfile.ZipFile(path) as archive:
            broken = archive.testzip()
            if broken is not None:
                return False, f"Arsip dokumen rusak pada entri '{broken}'."
            names = set(archive.namelist())
            for required in REQUIRED_ZIP_ENTRIES:
                if required not in names:
                    return False, (
                        "Struktur dokumen tidak valid. "
                        f"Entri wajib '{required}' tidak ditemukan."
                    )
    except zipfile.BadZipFile:
        return False, (
            "File bukan dokumen .docx yang valid. "
            "Format lama .doc tidak didukung, simpan ulang sebagai .docx."
        )
    except PermissionError:
        return False, "Akses ke file ditolak oleh sistem operasi."
    except OSError as exc:
        return False, f"File tidak dapat dibaca: {exc.strerror or exc}"
    return True, ""


def file_is_readable(path: Path) -> Tuple[bool, str]:
    """Check that the file can be opened for reading right now."""
    try:
        with path.open("rb") as handle:
            handle.read(1)
    except PermissionError:
        return False, (
            "File sedang digunakan oleh aplikasi lain atau akses ditolak. "
            "Tutup file tersebut di Microsoft Word lalu coba lagi."
        )
    except FileNotFoundError:
        return False, "File tidak ditemukan. Kemungkinan telah dipindahkan atau dihapus."
    except OSError as exc:
        return False, f"File tidak dapat dibuka: {exc.strerror or exc}"
    return True, ""


def directory_is_writable(directory: Path) -> Tuple[bool, str]:
    """Verify write permission by creating and removing a probe file."""
    try:
        if not directory.is_dir():
            return False, "Direktori tujuan tidak ditemukan."
        probe = directory / ".wordsplitter_write_probe.tmp"
        with probe.open("wb") as handle:
            handle.write(b"0")
        probe.unlink(missing_ok=True)
    except PermissionError:
        return False, (
            "Direktori tujuan bersifat read-only atau akses ditolak. "
            "Pindahkan dokumen ke folder lain lalu coba lagi."
        )
    except OSError as exc:
        return False, f"Direktori tujuan tidak dapat ditulis: {exc.strerror or exc}"
    return True, ""


def free_space_bytes(directory: Path) -> Optional[int]:
    """Return free space on the volume hosting the directory."""
    try:
        return shutil.disk_usage(directory).free
    except OSError:
        return None


def ensure_docx_extension(name: str) -> str:
    """Append the .docx extension when the user omitted or mistyped it."""
    cleaned = name.strip()
    if cleaned.lower().endswith(DOCX_EXTENSION):
        return cleaned[: -len(DOCX_EXTENSION)] + DOCX_EXTENSION
    return cleaned + DOCX_EXTENSION


def temp_sibling_path(target: Path, tag: str) -> Path:
    """Build a unique temporary path on the same volume as the final target.

    Staying on the same volume is what makes the final commit an atomic
    os.replace instead of a copy that can fail halfway through.
    """
    counter = 0
    while True:
        suffix = "" if counter == 0 else f"_{counter}"
        candidate = target.parent / f"~wordsplitter_{tag}{suffix}_{os.getpid()}.tmp.docx"
        if not candidate.exists():
            return candidate
        counter += 1


def safe_unlink(path: Optional[Path]) -> None:
    """Delete a file, ignoring every failure. Used only for temporary artefacts."""
    if path is None:
        return
    try:
        path.unlink(missing_ok=True)
    except Exception:  # noqa: BLE001 - cleanup must never raise
        pass


def word_owner_files(path: Path) -> list[Path]:
    """Return the possible Word lock file names for a given document.

    Word creates an owner file named "~$" followed by the document name with its
    first two characters removed when the name is long enough. Both variants are
    returned so that a crashed session does not leave debris behind.
    """
    name = path.name
    candidates = {f"~${name}"}
    if len(name) > 2:
        candidates.add(f"~${name[2:]}")
    return [path.parent / candidate for candidate in candidates]


def cleanup_temp_artifact(path: Optional[Path]) -> None:
    """Remove a temporary document together with any Word lock file."""
    if path is None:
        return
    safe_unlink(path)
    for owner in word_owner_files(path):
        safe_unlink(owner)


def files_are_identical(first: Path, second: Path) -> bool:
    """Byte comparison used to detect a split that produced two equal outputs."""
    try:
        if first.stat().st_size != second.stat().st_size:
            return False
        with first.open("rb") as fa, second.open("rb") as fb:
            while True:
                chunk_a = fa.read(65536)
                chunk_b = fb.read(65536)
                if chunk_a != chunk_b:
                    return False
                if not chunk_a:
                    return True
    except OSError:
        return False

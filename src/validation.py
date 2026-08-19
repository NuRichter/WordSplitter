"""Validation rules for input files, page selection and output file names.

Every function here is pure and deterministic, which makes the rules auditable
and testable without Microsoft Word.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

from utils import (
    DOCX_EXTENSION,
    ensure_docx_extension,
    file_is_readable,
    is_valid_docx_container,
)

WINDOWS_INVALID_CHARS = set('\\/:*?"<>|')
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
MAX_STEM_LENGTH = 180
_CONTROL_CHARS = re.compile(r"[\x00-\x1f\x7f]")


class ValidationError(Exception):
    """Raised when user supplied data violates a documented rule."""


@dataclass(frozen=True)
class SplitPlan:
    """Resolved split geometry.

    part_one covers pages 1 .. split_page - 1.
    part_two covers pages split_page .. total_pages.
    """

    page_a: int
    page_b: int
    split_page: int
    total_pages: int

    @property
    def description(self) -> str:
        return (
            f"Part 1 = halaman 1-{self.split_page - 1}, "
            f"Part 2 = halaman {self.split_page}-{self.total_pages}"
        )


def validate_input_file(raw_path: str) -> Path:
    """Validate the selected document and return its resolved path."""
    if not raw_path or not raw_path.strip():
        raise ValidationError("Belum ada file yang dipilih.")

    path = Path(raw_path).expanduser()
    try:
        path = path.resolve(strict=False)
    except OSError as exc:
        raise ValidationError(f"Path file tidak valid: {exc}") from exc

    if path.suffix.lower() != DOCX_EXTENSION:
        raise ValidationError(
            "Hanya file dengan ekstensi .docx yang didukung. "
            "File .doc, .rtf, .odt dan .pdf tidak dapat diproses."
        )
    if not path.exists():
        raise ValidationError("File tidak ditemukan. Kemungkinan telah dipindahkan atau dihapus.")
    if not path.is_file():
        raise ValidationError("Path yang dipilih bukan sebuah file.")

    readable, reason = file_is_readable(path)
    if not readable:
        raise ValidationError(reason)

    valid, reason = is_valid_docx_container(path)
    if not valid:
        raise ValidationError(reason)

    return path


def parse_page_number(raw: str, label: str) -> int:
    """Convert user text into a positive page index."""
    text = (raw or "").strip()
    if not text:
        raise ValidationError(f"{label} belum diisi.")
    if not re.fullmatch(r"[+-]?\d+", text):
        raise ValidationError(f"{label} harus berupa angka bulat.")
    value = int(text)
    if value < 0:
        raise ValidationError(f"{label} tidak boleh bernilai negatif.")
    if value == 0:
        raise ValidationError(f"{label} tidak boleh bernilai 0. Penomoran halaman dimulai dari 1.")
    return value


def compute_split_plan(page_a: int, page_b: int, total_pages: int) -> SplitPlan:
    """Resolve the split boundary from the two pages selected by the user.

    The boundary is placed at the midpoint between Page A and Page B. With
    A = 10 and B = 11 the midpoint is 10.5, so the split falls between page 10
    and page 11 and the first page of Part 2 is page 11. With A = 10 and B = 12
    the midpoint is 11, and the convention adopted here keeps the midpoint page
    inside Part 1, so the split falls between page 11 and page 12.

    General rule: split_page = floor((A + B) / 2) + 1.
    """
    if total_pages < 2:
        raise ValidationError(
            "Dokumen hanya memiliki satu halaman sehingga tidak dapat dipisah."
        )
    if page_a == page_b:
        raise ValidationError("Page A dan Page B tidak boleh sama.")

    low, high = (page_a, page_b) if page_a < page_b else (page_b, page_a)

    for label, value in (("Page A", page_a), ("Page B", page_b)):
        if value > total_pages:
            raise ValidationError(
                f"{label} = {value} berada di luar rentang dokumen. "
                f"Dokumen memiliki {total_pages} halaman."
            )

    split_page = (low + high) // 2 + 1
    if split_page < 2 or split_page > total_pages:
        raise ValidationError(
            "Titik pisah yang dihitung berada di luar rentang dokumen. "
            "Pilih kombinasi halaman lain."
        )

    return SplitPlan(
        page_a=page_a, page_b=page_b, split_page=split_page, total_pages=total_pages
    )


def validate_output_name(raw: str, label: str) -> str:
    """Validate a file name against Windows naming rules and normalise it."""
    text = (raw or "").strip()
    if not text:
        raise ValidationError(f"{label} tidak boleh kosong.")
    if _CONTROL_CHARS.search(text):
        raise ValidationError(f"{label} mengandung karakter kontrol yang tidak diizinkan.")

    offending = sorted({ch for ch in text if ch in WINDOWS_INVALID_CHARS})
    if offending:
        raise ValidationError(
            f"{label} mengandung karakter yang dilarang Windows: "
            + " ".join(offending)
        )

    name = ensure_docx_extension(text)
    stem = name[: -len(DOCX_EXTENSION)]

    if not stem.strip():
        raise ValidationError(f"{label} tidak boleh hanya berisi ekstensi.")
    if stem != stem.rstrip(" ."):
        raise ValidationError(f"{label} tidak boleh diakhiri spasi atau titik.")
    if stem.upper() in WINDOWS_RESERVED_NAMES:
        raise ValidationError(
            f"{label} menggunakan nama yang dicadangkan Windows: {stem}."
        )
    if len(stem) > MAX_STEM_LENGTH:
        raise ValidationError(
            f"{label} terlalu panjang. Maksimum {MAX_STEM_LENGTH} karakter."
        )
    return name


def resolve_output_paths(
    source: Path, name_one: str, name_two: str
) -> Tuple[Path, Path]:
    """Validate both names and resolve them inside the source directory."""
    first = validate_output_name(name_one, "File 1 Name")
    second = validate_output_name(name_two, "File 2 Name")

    directory = source.parent
    path_one = directory / first
    path_two = directory / second

    if str(path_one).lower() == str(path_two).lower():
        raise ValidationError(
            "File 1 dan File 2 menghasilkan path output yang sama. "
            "Gunakan nama yang berbeda."
        )
    for path, label in ((path_one, "File 1"), (path_two, "File 2")):
        if str(path).lower() == str(source).lower():
            raise ValidationError(
                f"{label} memiliki nama yang sama dengan file asli. "
                "Hal ini akan menimpa dokumen sumber, gunakan nama lain."
            )
        if len(str(path)) > 255:
            raise ValidationError(
                f"Path output untuk {label} melebihi batas panjang yang aman."
            )
    return path_one, path_two


def existing_outputs(path_one: Path, path_two: Path) -> list[Path]:
    """Return the output paths that already exist on disk."""
    return [path for path in (path_one, path_two) if path.exists()]


def estimate_required_space(source_size: int) -> int:
    """Conservative estimate of the disk space needed by the whole operation.

    Two temporary copies plus two final files, plus a safety margin.
    """
    return source_size * 4 + 16 * 1024 * 1024


def describe_page_selection(page_a: Optional[int], page_b: Optional[int]) -> str:
    if page_a is None or page_b is None:
        return "belum ditentukan"
    return f"Page A = {page_a}, Page B = {page_b}"

"""Validation rules for input files, page selection and output file names.

Every function here is pure and deterministic, which makes the rules auditable
and testable without Microsoft Word. Failures raise ValidationError with a
locale key; parameters named "label" hold a locale key themselves and are
translated by the GUI before rendering.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Tuple

from errors import AppError, ValidationError
from utils import (
    DOCX_EXTENSION,
    WORD_MAX_PATH,
    check_docx_container,
    check_readable,
    ensure_docx_extension,
)

WINDOWS_INVALID_CHARS = set('\\/:*?"<>|')
WINDOWS_RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
    "COM¹", "COM²", "COM³", "LPT¹", "LPT²", "LPT³",
}
MAX_STEM_LENGTH = 180
_CONTROL_CHARS = re.compile(r"[\x00-\x1f\x7f]")


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
    def part_one(self) -> Tuple[int, int]:
        return 1, self.split_page - 1

    @property
    def part_two(self) -> Tuple[int, int]:
        return self.split_page, self.total_pages

    @property
    def description(self) -> str:
        return (
            f"Part 1 = pages 1-{self.split_page - 1}, "
            f"Part 2 = pages {self.split_page}-{self.total_pages}"
        )


def validate_input_file(raw_path: str) -> Path:
    """Validate the selected document and return its resolved path."""
    if not raw_path or not str(raw_path).strip():
        raise ValidationError("err.no_file")

    path = Path(raw_path).expanduser()
    try:
        path = path.resolve(strict=False)
    except (OSError, RuntimeError) as exc:
        raise ValidationError("err.not_found", detail=str(exc)) from exc

    if path.suffix.lower() != DOCX_EXTENSION:
        raise ValidationError("err.not_docx_ext")
    if not path.exists():
        raise ValidationError("err.not_found")
    if not path.is_file():
        raise ValidationError("err.not_a_file")

    try:
        check_readable(path)
        check_docx_container(path)
    except AppError as exc:
        raise ValidationError(exc.key, exc.detail, **exc.params) from None
    return path


def parse_page_number(raw: str, label: str) -> int:
    """Convert user text into a positive page index.

    Any Unicode decimal digits are accepted (Arabic-Indic, Persian, Devanagari,
    full width), because people type numbers with the keyboard of their
    language. Python's int() maps all of them to the same value.
    """
    text = unicodedata.normalize("NFKC", raw or "").strip()
    if not text:
        raise ValidationError("err.page_empty", label=label)
    if not re.fullmatch(r"[+-]?\d+", text):
        raise ValidationError("err.page_not_int", label=label)
    value = int(text)
    if value < 1:
        raise ValidationError("err.page_zero", label=label)
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
        raise ValidationError("err.single_page")
    for label, value in (("label.page_a", page_a), ("label.page_b", page_b)):
        if value > total_pages:
            raise ValidationError("err.page_out_of_range", label=label, value=value, total=total_pages)
    if page_a == page_b:
        raise ValidationError("err.page_same")

    low, high = (page_a, page_b) if page_a < page_b else (page_b, page_a)
    split_page = (low + high) // 2 + 1
    if split_page < 2 or split_page > total_pages:
        raise ValidationError("err.split_out_of_range")
    return SplitPlan(page_a=page_a, page_b=page_b, split_page=split_page, total_pages=total_pages)


def validate_output_name(raw: str, label: str) -> str:
    """Validate a file name against Windows naming rules and normalise it."""
    text = unicodedata.normalize("NFC", raw or "").strip()
    if not text:
        raise ValidationError("err.name_empty", label=label)
    if _CONTROL_CHARS.search(text):
        raise ValidationError("err.name_control", label=label)

    offending = sorted({ch for ch in text if ch in WINDOWS_INVALID_CHARS})
    if offending:
        raise ValidationError("err.name_invalid_chars", label=label, chars=" ".join(offending))

    name = ensure_docx_extension(text)
    stem = name[: -len(DOCX_EXTENSION)]

    if not stem.strip():
        raise ValidationError("err.name_only_ext", label=label)
    if stem != stem.rstrip(" ."):
        raise ValidationError("err.name_trailing", label=label)
    if stem.split(".")[0].strip().upper() in WINDOWS_RESERVED_NAMES:
        raise ValidationError("err.name_reserved", label=label, name=stem)
    if len(stem) > MAX_STEM_LENGTH:
        raise ValidationError("err.name_too_long", label=label, max=MAX_STEM_LENGTH)
    return name


def resolve_output_paths(source: Path, directory: Path, name_one: str, name_two: str) -> Tuple[Path, Path]:
    """Validate both names and resolve them inside the destination directory."""
    first = validate_output_name(name_one, "label.name1")
    second = validate_output_name(name_two, "label.name2")

    path_one = directory / first
    path_two = directory / second

    if str(path_one).casefold() == str(path_two).casefold():
        raise ValidationError("err.names_same")
    for path, label in ((path_one, "label.name1"), (path_two, "label.name2")):
        if str(path).casefold() == str(source).casefold():
            raise ValidationError("err.name_is_source", label=label)
        if len(str(path)) > WORD_MAX_PATH:
            raise ValidationError("err.output_path_too_long")
    return path_one, path_two


def existing_outputs(path_one: Path, path_two: Path) -> list:
    """Return the output paths that already exist on disk."""
    return [path for path in (path_one, path_two) if path.exists()]


def estimate_required_space(source_size: int) -> int:
    """Conservative estimate of the disk space needed by the whole operation.

    One working copy, two parts and two staged final files, plus a margin.
    """
    return source_size * 5 + 16 * 1024 * 1024

"""Interface languages.

Only two locale files are ever loaded: English, which every other language
falls back to key by key, and the language in use. Each language also declares
its writing direction and the Windows UI font that renders its script best,
so labels never fall back to boxes or to a mismatched font.
"""

from __future__ import annotations

import json
import string
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Optional

import utils


@dataclass(frozen=True)
class Language:
    code: str
    native: str
    english: str
    font: str = "Segoe UI"
    rtl: bool = False


# Thirty widely spoken languages, listed by their own names.
LANGUAGES = (
    Language("en", "English", "English"),
    Language("id", "Bahasa Indonesia", "Indonesian"),
    Language("ms", "Bahasa Melayu", "Malay"),
    Language("jv", "Basa Jawa", "Javanese"),
    Language("de", "Deutsch", "German"),
    Language("es", "Español", "Spanish"),
    Language("fil", "Filipino", "Filipino"),
    Language("fr", "Français", "French"),
    Language("it", "Italiano", "Italian"),
    Language("sw", "Kiswahili", "Swahili"),
    Language("nl", "Nederlands", "Dutch"),
    Language("pl", "Polski", "Polish"),
    Language("pt-BR", "Português (Brasil)", "Portuguese (Brazil)"),
    Language("vi", "Tiếng Việt", "Vietnamese"),
    Language("tr", "Türkçe", "Turkish"),
    Language("ru", "Русский", "Russian"),
    Language("uk", "Українська", "Ukrainian"),
    Language("ar", "العربية", "Arabic", rtl=True),
    Language("fa", "فارسی", "Persian", rtl=True),
    Language("ur", "اردو", "Urdu", rtl=True),
    Language("hi", "हिन्दी", "Hindi", font="Nirmala UI"),
    Language("mr", "मराठी", "Marathi", font="Nirmala UI"),
    Language("bn", "বাংলা", "Bengali", font="Nirmala UI"),
    Language("ta", "தமிழ்", "Tamil", font="Nirmala UI"),
    Language("te", "తెలుగు", "Telugu", font="Nirmala UI"),
    Language("th", "ไทย", "Thai", font="Leelawadee UI"),
    Language("zh-CN", "简体中文", "Chinese (Simplified)", font="Microsoft YaHei UI"),
    Language("zh-TW", "繁體中文", "Chinese (Traditional)", font="Microsoft JhengHei UI"),
    Language("ja", "日本語", "Japanese", font="Yu Gothic UI"),
    Language("ko", "한국어", "Korean", font="Malgun Gothic"),
)
BY_CODE = {language.code: language for language in LANGUAGES}
DEFAULT = "en"

# Windows language tags that do not map onto a code by their prefix alone.
_TAG_ALIASES = {
    "zh-tw": "zh-TW", "zh-hk": "zh-TW", "zh-mo": "zh-TW", "zh-hant": "zh-TW",
    "zh-cn": "zh-CN", "zh-sg": "zh-CN", "zh-hans": "zh-CN",
    "pt": "pt-BR", "tl": "fil", "fil": "fil", "jv": "jv", "jv-latn": "jv",
}


def locales_dir() -> Path:
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / "locales"


def match_tag(tag: Optional[str]) -> str:
    """Map a BCP 47 tag such as 'id-ID' or 'zh-Hant-TW' onto a supported code."""
    if not tag:
        return DEFAULT
    lowered = tag.replace("_", "-").lower()
    parts = lowered.split("-")
    candidates = []
    if len(parts) >= 3:
        candidates.append(f"{parts[0]}-{parts[-1]}")
        candidates.append(f"{parts[0]}-{parts[1]}")
    if len(parts) >= 2:
        candidates.append(f"{parts[0]}-{parts[1]}")
    candidates.append(parts[0])
    for candidate in candidates:
        if candidate in _TAG_ALIASES:
            return _TAG_ALIASES[candidate]
        for language in LANGUAGES:
            if language.code.lower() == candidate:
                return language.code
    if parts[0] == "zh":
        return "zh-CN"
    return DEFAULT


def _load(code: str) -> Dict[str, str]:
    path = locales_dir() / f"{code}.json"
    try:
        with path.open("r", encoding="utf-8") as handle:
            data = json.load(handle)
        return {str(k): str(v) for k, v in data.items() if isinstance(v, str)}
    except (OSError, ValueError):
        return {}


class _SafeDict(dict):
    def __missing__(self, key: str) -> str:
        return "{" + key + "}"


class Translator:
    """Lookups for one language with a per key English fallback."""

    def __init__(self, code: str) -> None:
        self.language = BY_CODE.get(code, BY_CODE[DEFAULT])
        self.code = self.language.code
        self._fallback = _load(DEFAULT)
        self._strings = self._fallback if self.code == DEFAULT else _load(self.code)

    @property
    def rtl(self) -> bool:
        return self.language.rtl

    def has(self, key: str) -> bool:
        return key in self._strings or key in self._fallback

    def t(self, key: str, **params: Any) -> str:
        text = self._strings.get(key) or self._fallback.get(key) or key
        if not params:
            return text
        prepared = {}
        for name, value in params.items():
            if name == "label" and isinstance(value, str) and self.has(value):
                value = self.t(value)
            elif name.endswith("_bytes") and isinstance(value, (int, float)):
                value = self.size(int(value))
            prepared[name] = value
        try:
            return string.Formatter().vformat(text, (), _SafeDict(prepared))
        except (ValueError, IndexError):
            return text

    def size(self, size: int) -> str:
        units = tuple(self.t(f"unit.{u}") for u in ("b", "kb", "mb", "gb")) + ("TB",)
        return utils.format_size(size, units, self.t("num.decimal") or ".")

    def message(self, error: Dict[str, Any]) -> str:
        """Render an error or warning dictionary {key, params}."""
        return self.t(str(error.get("key") or "err.unexpected"), **(error.get("params") or {}))

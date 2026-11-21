"""Small persistent preferences (language, theme, last folder).

Stored as JSON next to the logs. Reads tolerate a missing or damaged file and
writes are atomic, so a power cut can never leave a half written settings file
that breaks the next start.
"""

from __future__ import annotations

import json
import os
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Optional

import logger as app_logger

THEMES = ("light", "dark")


@dataclass
class Settings:
    language: Optional[str] = None
    theme: Optional[str] = None
    last_folder: Optional[str] = None

    @staticmethod
    def path() -> Path:
        return app_logger.app_data_dir() / "settings.json"

    @classmethod
    def load(cls) -> "Settings":
        try:
            with cls.path().open("r", encoding="utf-8") as handle:
                data = json.load(handle)
        except (OSError, ValueError):
            return cls()
        if not isinstance(data, dict):
            return cls()
        settings = cls()
        for name in ("language", "theme", "last_folder"):
            value = data.get(name)
            if isinstance(value, str) and value:
                setattr(settings, name, value)
        if settings.theme not in THEMES:
            settings.theme = None
        return settings

    def save(self) -> None:
        path = self.path()
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            staging = path.with_suffix(".tmp")
            with staging.open("w", encoding="utf-8") as handle:
                json.dump(asdict(self), handle, ensure_ascii=False, indent=2)
            os.replace(staging, path)
        except OSError:
            pass  # preferences are a convenience, never a reason to fail

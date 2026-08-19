"""Logging subsystem for WordSplitter.

Logs are written to a rotating file inside the per-user application data
directory. Logging must never be able to break the application, therefore every
handler installation is guarded and falls back to a temporary directory, and
finally to a no-op handler.
"""

from __future__ import annotations

import logging
import logging.handlers
import os
import sys
import tempfile
from pathlib import Path
from typing import Optional

APP_NAME = "WordSplitter"
LOG_FILE_NAME = "wordsplitter.log"
MAX_BYTES = 1_000_000
BACKUP_COUNT = 3

_LOG_PATH: Optional[Path] = None


def app_data_dir() -> Path:
    """Return the writable per-user data directory for the application."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    if not base:
        base = tempfile.gettempdir()
    return Path(base) / APP_NAME


def log_file_path() -> Optional[Path]:
    """Return the path of the active log file, or None if logging is disabled."""
    return _LOG_PATH


def _build_file_handler(directory: Path) -> logging.Handler:
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / LOG_FILE_NAME
    handler = logging.handlers.RotatingFileHandler(
        path, maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT, encoding="utf-8"
    )
    handler.setFormatter(
        logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(threadName)-12s | %(name)s | %(message)s"
        )
    )
    global _LOG_PATH
    _LOG_PATH = path
    return handler


def setup_logging(level: int = logging.INFO) -> logging.Logger:
    """Configure the root logger. Safe to call more than once."""
    root = logging.getLogger()
    if getattr(root, "_wordsplitter_configured", False):
        return logging.getLogger(APP_NAME)

    root.setLevel(level)

    handler: Optional[logging.Handler] = None
    for candidate in (app_data_dir() / "logs", Path(tempfile.gettempdir()) / APP_NAME):
        try:
            handler = _build_file_handler(candidate)
            break
        except Exception:  # noqa: BLE001 - logging must never abort startup
            handler = None

    if handler is None:
        handler = logging.NullHandler()

    root.addHandler(handler)

    # A console handler is only useful for the non frozen developer run.
    if not getattr(sys, "frozen", False):
        console = logging.StreamHandler(stream=sys.stderr)
        console.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        root.addHandler(console)

    root._wordsplitter_configured = True  # type: ignore[attr-defined]
    logger = logging.getLogger(APP_NAME)
    logger.info("Logging initialised. Log file: %s", _LOG_PATH)
    return logger


def get_logger(name: str = APP_NAME) -> logging.Logger:
    return logging.getLogger(name)

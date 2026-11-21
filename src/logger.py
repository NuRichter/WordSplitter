"""Logging subsystem for WordSplitter.

Logs are written to a rotating file inside the per-user application data
directory. The GUI and the worker process each own their own file, so the two
never fight over a rotation. Logging must never be able to break the
application: every handler installation is guarded and falls back to the temp
directory, and finally to a no-op handler.
"""

from __future__ import annotations

import logging
import logging.handlers
import os
import sys
import tempfile
from pathlib import Path
from typing import Optional

from app_info import APP_NAME

GUI_LOG = "wordsplitter.log"
WORKER_LOG = "worker.log"
MAX_BYTES = 1_000_000
BACKUP_COUNT = 3

_LOG_PATH: Optional[Path] = None


def app_data_dir() -> Path:
    """Return the writable per-user data directory for the application."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA") or tempfile.gettempdir()
    return Path(base) / APP_NAME


def log_file_path() -> Optional[Path]:
    """Return the path of the active log file, or None if logging is disabled."""
    return _LOG_PATH


def _build_file_handler(directory: Path, file_name: str) -> logging.Handler:
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / file_name
    handler = logging.handlers.RotatingFileHandler(
        path, maxBytes=MAX_BYTES, backupCount=BACKUP_COUNT, encoding="utf-8", delay=True
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s | %(levelname)-7s | %(process)d | %(threadName)s | %(name)s | %(message)s")
    )
    global _LOG_PATH
    _LOG_PATH = path
    return handler


def setup_logging(file_name: str = GUI_LOG, level: int = logging.INFO) -> logging.Logger:
    """Configure the root logger once per process."""
    root = logging.getLogger()
    if getattr(root, "_wordsplitter_configured", False):
        return logging.getLogger(APP_NAME)
    root.setLevel(level)

    handler: Optional[logging.Handler] = None
    for candidate in (app_data_dir() / "logs", Path(tempfile.gettempdir()) / APP_NAME / "logs"):
        try:
            handler = _build_file_handler(candidate, file_name)
            break
        except Exception:  # noqa: BLE001 - logging must never abort startup
            handler = None
    root.addHandler(handler or logging.NullHandler())

    # A console handler only helps a developer running from source with a console.
    if not getattr(sys, "frozen", False) and sys.stderr is not None and file_name == GUI_LOG:
        console = logging.StreamHandler(stream=sys.stderr)
        console.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))
        root.addHandler(console)

    logging.raiseExceptions = False  # a broken log file must stay silent
    root._wordsplitter_configured = True  # type: ignore[attr-defined]
    return logging.getLogger(APP_NAME)


def get_logger(name: str = APP_NAME) -> logging.Logger:
    return logging.getLogger(name)

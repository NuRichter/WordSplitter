"""Worker process entry point.

The GUI never loads pywin32 or talks to Word. For every analysis or split it
starts a short lived copy of itself with `--worker <request.json>`. The worker
does the COM work, streams JSON events on its standard output and exits.

This buys three guarantees:
* the GUI's idle footprint never includes COM, pywin32 or a Word instance;
* a crash inside Word or COM can only end the worker, never the window;
* Cancel is immediate: the GUI terminates the worker, the kernel closes the
  worker's kill-on-close job, and Word goes down with it.
"""

from __future__ import annotations

import json
import logging
import os
import sys
import threading
import time
from pathlib import Path
from typing import Any, Optional

import logger as app_logger
import utils
import winsys
from errors import AppError, OperationCancelled

LOGGER = logging.getLogger("WordSplitter.worker")


class EventStream:
    """Line oriented JSON writer bound to the inherited stdout pipe."""

    def __init__(self) -> None:
        self._out = self._open()

    @staticmethod
    def _open():
        if winsys.IS_WINDOWS:
            try:
                import msvcrt

                handle = winsys.std_output_handle()
                if handle:
                    fd = msvcrt.open_osfhandle(handle, os.O_WRONLY)
                    return os.fdopen(fd, "wb", buffering=0)
            except Exception:  # noqa: BLE001
                pass
        stream = getattr(sys.stdout, "buffer", None)
        return stream

    def emit(self, kind: str, **data: Any) -> None:
        payload = {"t": kind, **data}
        line = json.dumps(payload, ensure_ascii=True, default=str) + "\n"
        if self._out is None:
            LOGGER.info("event %s", line.strip())
            return
        try:
            self._out.write(line.encode("ascii"))
            self._out.flush()
        except OSError:
            # The GUI is gone; nothing left to report to.
            self._out = None


HARD_STOP_SECONDS = 60


def _watch_parent(cancel: threading.Event) -> None:
    """Treat "cancel" on stdin, or the pipe closing, as a request to stop.

    The pipe closes by itself when the GUI exits or is killed, so a worker can
    never outlive its window. The worker stops Word on its own (it is the
    only one that knows Word's pid, even while Word is still starting), and
    only if it is stuck does the timer below end the process outright.
    """
    handle = winsys.std_input_handle()
    if not handle:
        return
    try:
        import msvcrt

        stream = os.fdopen(msvcrt.open_osfhandle(handle, os.O_RDONLY), "rb", buffering=0)
        while True:
            line = stream.readline()
            if not line or line.strip() == b"cancel":
                break
    except Exception:  # noqa: BLE001
        pass
    LOGGER.info("Cancel requested.")
    cancel.set()
    import word_engine  # cheap: pywin32 is only loaded when Word is started

    word_engine.WordSession.cancel_all()
    deadline = time.monotonic() + HARD_STOP_SECONDS
    while time.monotonic() < deadline:
        time.sleep(0.5)
    engine_module = sys.modules.get("split_engine")
    if engine_module is not None and engine_module.COMMITTING.is_set():
        return  # never cut a commit short
    LOGGER.error("Worker did not stop after cancel; exiting hard.")
    os._exit(3)


def run(request_path: str) -> int:
    app_logger.setup_logging(app_logger.WORKER_LOG)
    events = EventStream()
    job = winsys.KillOnCloseJob()
    job_dir: Optional[Path] = None
    started = time.perf_counter()
    cancel = threading.Event()
    threading.Thread(target=_watch_parent, args=(cancel,), name="ParentWatch", daemon=True).start()
    try:
        request = json.loads(Path(request_path).read_text(encoding="utf-8"))
        job_dir = Path(request["job_dir"])
        kind = request.get("kind")
        LOGGER.info("Worker %s started for %s", kind, request.get("source"))

        import split_engine  # imported here so a bad request fails fast and cheap

        if kind == "analyze":
            result = split_engine.analyze(Path(request["source"]), job_dir, events.emit, job, cancel=cancel)
            events.emit("done", result=result)
        elif kind == "split":
            outcome = split_engine.split(split_engine.SplitRequest.from_dict(request), events.emit, job, cancel=cancel)
            events.emit("done", result=outcome.to_dict())
        else:
            raise AppError("err.unexpected", detail=f"unknown job kind {kind!r}")
    except OperationCancelled:
        LOGGER.info("Worker cancelled.")
        events.emit("cancelled")
    except AppError as exc:
        LOGGER.warning("Worker failed: %s", exc)
        events.emit("error", error=exc.to_dict())
    except Exception as exc:  # noqa: BLE001
        LOGGER.exception("Unexpected worker failure")
        events.emit("error", error=AppError("err.unexpected", detail=f"{type(exc).__name__}: {exc}").to_dict())
    finally:
        engine = sys.modules.get("word_engine")
        if engine is not None:
            engine.settle_exits()
        utils.remove_tree(job_dir, attempts=10)
        LOGGER.info("Worker finished in %.2f s", time.perf_counter() - started)
    return 0

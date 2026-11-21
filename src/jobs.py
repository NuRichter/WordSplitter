"""GUI side of the worker protocol.

A WorkerJob starts the worker process, turns its JSON lines into queue items
for the Tk thread, and offers an immediate, reliable cancel. Nothing in here
blocks the Tk event loop.
"""

from __future__ import annotations

import json
import logging
import os
import queue
import subprocess
import sys
import threading
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

import utils
import winsys

LOGGER = logging.getLogger("WordSplitter.jobs")

CREATE_NO_WINDOW = 0x08000000
_STALE_SWEEP_DONE = False


def worker_command(request_path: Path) -> list:
    if getattr(sys, "frozen", False):
        return [sys.executable, "--worker", str(request_path)]
    main_py = Path(__file__).resolve().parent / "main.py"
    return [sys.executable, str(main_py), "--worker", str(request_path)]


def sweep_stale_jobs() -> None:
    """Remove job folders left behind by a crash or a power loss (once per run)."""
    global _STALE_SWEEP_DONE
    if _STALE_SWEEP_DONE:
        return
    _STALE_SWEEP_DONE = True

    def _sweep() -> None:
        root = utils.jobs_root()
        try:
            entries = list(root.iterdir())
        except OSError:
            return
        for entry in entries:
            parts = entry.name.split("-")
            if len(parts) < 3 or parts[0] != "job" or not parts[1].isdigit():
                continue
            pid = int(parts[1])
            if pid == os.getpid() or winsys.pid_is_alive(pid):
                continue
            utils.remove_tree(entry, attempts=1)

    threading.Thread(target=_sweep, name="StaleSweep", daemon=True).start()


class WorkerJob:
    """One analysis or split running in a separate process."""

    def __init__(self, request: Dict[str, Any]) -> None:
        sweep_stale_jobs()
        self.kind = request.get("kind")
        self.job_dir = utils.new_job_dir()
        self.events: "queue.Queue[Tuple[str, Any]]" = queue.Queue()
        self.committing = False
        self.cancelled = False
        self.final_seen = False
        self._finished = False
        request = dict(request, job_dir=str(self.job_dir))
        request_path = self.job_dir / "request.json"
        request_path.write_text(json.dumps(request, ensure_ascii=False), encoding="utf-8")
        self._process = subprocess.Popen(
            worker_command(request_path),
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_NO_WINDOW if winsys.IS_WINDOWS else 0,
            close_fds=True,
        )
        # The worker is deliberately not tied to a kill-on-close job: when this
        # process dies, the worker's stdin pipe closes and the worker shuts
        # Word down itself, which also covers a Word that is still starting.
        self._reader = threading.Thread(target=self._read, name=f"{self.kind}-reader", daemon=True)
        self._reader.start()
        LOGGER.info("Worker %s started (pid %s)", self.kind, self._process.pid)

    def _read(self) -> None:
        stream = self._process.stdout
        final_seen = False
        try:
            for raw in iter(stream.readline, b""):
                try:
                    event = json.loads(raw.decode("utf-8", "replace"))
                except ValueError:
                    continue
                kind = event.pop("t", "")
                if kind == "commit":
                    self.committing = True
                if kind in ("done", "error"):
                    final_seen = True
                    self.final_seen = True
                self.events.put((kind, event))
        except (OSError, ValueError):
            pass
        finally:
            code = self._process.wait()
            if self.cancelled:
                self.events.put(("cancelled", {}))
            elif not final_seen:
                LOGGER.error("Worker ended without a result (exit code %s)", code)
                self.events.put(("error", {"error": {"key": "err.worker_crashed", "params": {},
                                                     "detail": f"exit code {code}"}}))
            self.events.put(("exit", {"code": code}))
            self._cleanup()

    def cancel(self) -> bool:
        """Ask the worker to stop. Refused only once results are being committed."""
        if self.committing or self.final_seen or self._finished:
            return False
        self.cancelled = True
        try:
            self._process.stdin.write(b"cancel\n")
            self._process.stdin.flush()
        except (OSError, ValueError):
            pass
        threading.Thread(target=self._enforce_stop, name="CancelGuard", daemon=True).start()
        return True

    def _enforce_stop(self, grace: float = 20.0) -> None:
        """If the worker ignores the request, end it (its job object ends Word)."""
        try:
            self._process.wait(timeout=grace)
        except subprocess.TimeoutExpired:
            LOGGER.error("Worker ignored cancel; terminating it.")
            try:
                self._process.kill()
            except OSError:
                pass

    def _cleanup(self) -> None:
        self._finished = True
        for stream in (self._process.stdin, self._process.stdout):
            try:
                stream.close()
            except (OSError, ValueError, AttributeError):
                pass
        utils.remove_tree(self.job_dir)

    @property
    def running(self) -> bool:
        return not self._finished

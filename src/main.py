"""WordSplitter entry point.

Two roles live in one executable:

* no arguments: the window (Tk only, no COM);
* `--worker <request.json>`: a short lived background task that talks to
  Microsoft Word and reports back through its standard output.

The worker branch returns before tkinter is ever imported, and the GUI branch
never imports pywin32, which keeps each process as small as its job allows.
"""

from __future__ import annotations

import sys
from pathlib import Path

# Allow both `python src/main.py` and the PyInstaller bundle to resolve the
# sibling modules without a package installation.
sys.path.insert(0, str(Path(__file__).resolve().parent))


def _worker(argv: list) -> int:
    import worker

    if len(argv) < 1:
        return 2
    return worker.run(argv[0])


def _gui() -> int:
    import os
    import threading
    import traceback

    import logger as app_logger
    import winsys
    from app_info import APP_NAME, VERSION

    log = app_logger.setup_logging(app_logger.GUI_LOG)

    def fatal(name: str) -> None:
        path = app_logger.log_file_path()
        text = f"{APP_NAME} hit an unrecoverable error and has to close.\n\n{name}"
        if path:
            text += f"\n\nLog: {path}"
        winsys.message_box(text, APP_NAME)

    def hook(exc_type, exc_value, exc_tb) -> None:
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_tb)
            return
        log.critical("Unhandled exception:\n%s", "".join(traceback.format_exception(exc_type, exc_value, exc_tb)))
        fatal(exc_type.__name__)

    def thread_hook(args) -> None:
        log.critical(
            "Unhandled exception in thread %s:\n%s",
            args.thread.name if args.thread else "?",
            "".join(traceback.format_exception(args.exc_type, args.exc_value, args.exc_traceback)),
        )

    sys.excepthook = hook
    threading.excepthook = thread_hook

    if os.name != "nt":
        log.error("Unsupported platform: %s", sys.platform)
        fatal("WordSplitter needs Windows with Microsoft Word.")
        return 2

    winsys.enable_dpi_awareness()
    log.info("%s %s starting (Python %s)", APP_NAME, VERSION, sys.version.split()[0])
    try:
        import gui

        gui.run()
    except Exception:  # noqa: BLE001
        log.exception("Fatal error at application level")
        fatal("ApplicationStartupError")
        return 1
    log.info("%s closed.", APP_NAME)
    return 0


def _check(argv: list) -> int:
    """`--check <file.docx>`: analyse a document through a real worker, no window.

    Prints one JSON line (pages, timings, Word status) on standard output. It
    exercises exactly the path the window uses, so it doubles as a support
    diagnostic and as a smoke test for packaged builds.
    """
    import json
    import os
    import queue
    import time

    import jobs
    import logger as app_logger
    import winsys

    app_logger.setup_logging(app_logger.GUI_LOG)
    report = {"file": argv[0] if argv else None, "frozen": bool(getattr(sys, "frozen", False))}
    started = time.perf_counter()
    try:
        job = jobs.WorkerJob({"kind": "analyze", "source": argv[0]})
        while True:
            kind, data = job.events.get(timeout=600)
            if kind == "progress":
                report.setdefault("stages", []).append([data.get("key"), round(time.perf_counter() - started, 2)])
            elif kind in ("done", "error", "cancelled"):
                report["outcome"] = kind
                report.update(data)
                report["seconds"] = round(time.perf_counter() - started, 2)
                break
    except (IndexError, queue.Empty) as exc:
        report["outcome"] = "error"
        report["error"] = {"key": "err.unexpected", "detail": type(exc).__name__}
    line = (json.dumps(report, ensure_ascii=True) + "\n").encode("ascii")
    handle = winsys.std_output_handle()
    if handle:
        import msvcrt

        with os.fdopen(msvcrt.open_osfhandle(handle, os.O_WRONLY), "wb") as out:
            out.write(line)
    elif sys.stdout is not None:
        sys.stdout.write(line.decode("ascii"))
    return 0 if report.get("outcome") == "done" else 1


def main(argv: list) -> int:
    if len(argv) > 1 and argv[1] == "--worker":
        return _worker(argv[2:])
    if len(argv) > 1 and argv[1] == "--check":
        return _check(argv[2:])
    return _gui()


if __name__ == "__main__":
    sys.exit(main(sys.argv))

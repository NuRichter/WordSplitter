"""Tkinter user interface for WordSplitter.

The interface owns no business logic. It validates nothing by itself beyond
reading widget state, delegates every rule to `validation`, and delegates every
document operation to `split_engine`. All long running work is executed on a
worker thread and communicated back through a queue that the Tk event loop
drains, so the window never freezes.
"""

from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Optional, Tuple

import logger as app_logger
import split_engine
import utils
import validation
from split_engine import OperationCancelled, SplitError, SplitRequest, SplitResult
from validation import ValidationError

LOGGER = logging.getLogger("WordSplitter.gui")

APP_TITLE = "WordSplitter"
WINDOW_MIN_SIZE = (620, 640)
POLL_INTERVAL_MS = 100


class ConflictDialog(tk.Toplevel):
    """Modal dialog offering a safe resolution for existing output files."""

    def __init__(self, parent: tk.Misc, conflicts: list[Path]) -> None:
        super().__init__(parent)
        self.title("File already exists")
        self.resizable(False, False)
        self.transient(parent)
        self.choice: str = "cancel"

        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)

        ttk.Label(body, text="File already exists.", font=("Segoe UI", 10, "bold")).pack(
            anchor="w"
        )
        listing = "\n".join(f"\u2022 {path.name}" for path in conflicts)
        ttk.Label(
            body,
            text=(
                "File berikut sudah ada di folder tujuan:\n\n"
                f"{listing}\n\n"
                "Pilih tindakan yang diinginkan."
            ),
            justify="left",
            wraplength=420,
        ).pack(anchor="w", pady=(8, 16))

        buttons = ttk.Frame(body)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Cancel", command=lambda: self._choose("cancel")).pack(
            side="right"
        )
        ttk.Button(
            buttons, text="Choose Another Name", command=lambda: self._choose("rename")
        ).pack(side="right", padx=8)
        ttk.Button(buttons, text="Replace", command=lambda: self._choose("replace")).pack(
            side="right"
        )

        self.protocol("WM_DELETE_WINDOW", lambda: self._choose("cancel"))
        self.bind("<Escape>", lambda _event: self._choose("cancel"))
        self.update_idletasks()
        self._centre_on(parent)
        self.grab_set()

    def _centre_on(self, parent: tk.Misc) -> None:
        try:
            x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 3
            self.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except tk.TclError:
            pass

    def _choose(self, value: str) -> None:
        self.choice = value
        self.destroy()


class WordSplitterApp(ttk.Frame):
    """Main application frame."""

    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=16)
        self.master_window = master
        self.pack(fill="both", expand=True)

        self._queue: "queue.Queue[Tuple[str, Any]]" = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._cancel_event = threading.Event()
        self._busy = False

        self._source: Optional[Path] = None
        self._total_pages: Optional[int] = None

        self.var_file_name = tk.StringVar(value="Belum ada file dipilih.")
        self.var_file_location = tk.StringVar(value="-")
        self.var_file_size = tk.StringVar(value="-")
        self.var_file_status = tk.StringVar(value="-")
        self.var_pages = tk.StringVar(value="-")
        self.var_page_a = tk.StringVar()
        self.var_page_b = tk.StringVar()
        self.var_name_one = tk.StringVar(value="Part 1")
        self.var_name_two = tk.StringVar(value="Part 2")
        self.var_delete_original = tk.BooleanVar(value=False)
        self.var_status = tk.StringVar(value="Siap.")

        self._build_widgets()
        self._update_controls()
        self.after(POLL_INTERVAL_MS, self._drain_queue)

    # -- construction ------------------------------------------------------

    def _build_widgets(self) -> None:
        self.columnconfigure(0, weight=1)

        # File section
        file_frame = ttk.LabelFrame(self, text="1. Dokumen sumber", padding=12)
        file_frame.grid(row=0, column=0, sticky="ew")
        file_frame.columnconfigure(1, weight=1)

        self.btn_select = ttk.Button(
            file_frame, text="Select Word File", command=self._on_select_file
        )
        self.btn_select.grid(row=0, column=0, sticky="w")
        ttk.Label(file_frame, textvariable=self.var_file_name, wraplength=420).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )

        for row, (label, variable) in enumerate(
            (
                ("Lokasi", self.var_file_location),
                ("Ukuran", self.var_file_size),
                ("Status", self.var_file_status),
            ),
            start=1,
        ):
            ttk.Label(file_frame, text=f"{label}:").grid(
                row=row, column=0, sticky="w", pady=(6, 0)
            )
            ttk.Label(
                file_frame, textvariable=variable, wraplength=420, justify="left"
            ).grid(row=row, column=1, sticky="w", padx=(12, 0), pady=(6, 0))

        # Analysis section
        page_frame = ttk.LabelFrame(self, text="2. Halaman", padding=12)
        page_frame.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        page_frame.columnconfigure(3, weight=1)

        ttk.Label(page_frame, text="Jumlah halaman:").grid(row=0, column=0, sticky="w")
        ttk.Label(page_frame, textvariable=self.var_pages).grid(
            row=0, column=1, sticky="w", padx=(12, 0)
        )
        self.btn_analyze = ttk.Button(
            page_frame, text="Analisis Ulang", command=self._on_analyze
        )
        self.btn_analyze.grid(row=0, column=3, sticky="e")

        ttk.Label(page_frame, text="Page A:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.entry_page_a = ttk.Entry(page_frame, textvariable=self.var_page_a, width=8)
        self.entry_page_a.grid(row=1, column=1, sticky="w", padx=(12, 0), pady=(10, 0))

        ttk.Label(page_frame, text="Page B:").grid(row=2, column=0, sticky="w", pady=(6, 0))
        self.entry_page_b = ttk.Entry(page_frame, textvariable=self.var_page_b, width=8)
        self.entry_page_b.grid(row=2, column=1, sticky="w", padx=(12, 0), pady=(6, 0))

        ttk.Label(
            page_frame,
            text=(
                "Titik pisah berada di tengah antara Page A dan Page B. "
                "Contoh: Page A = 10 dan Page B = 11 memisahkan dokumen tepat "
                "pada batas halaman 10 dan 11."
            ),
            wraplength=520,
            justify="left",
            foreground="#555555",
        ).grid(row=3, column=0, columnspan=4, sticky="w", pady=(10, 0))

        # Output section
        out_frame = ttk.LabelFrame(self, text="3. Output", padding=12)
        out_frame.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        out_frame.columnconfigure(1, weight=1)

        ttk.Label(out_frame, text="File 1 Name:").grid(row=0, column=0, sticky="w")
        ttk.Entry(out_frame, textvariable=self.var_name_one).grid(
            row=0, column=1, sticky="ew", padx=(12, 0)
        )
        ttk.Label(out_frame, text="File 2 Name:").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(out_frame, textvariable=self.var_name_two).grid(
            row=1, column=1, sticky="ew", padx=(12, 0), pady=(6, 0)
        )
        ttk.Label(
            out_frame,
            text="Ekstensi .docx ditambahkan otomatis. File disimpan di folder yang sama dengan dokumen asli.",
            wraplength=520,
            justify="left",
            foreground="#555555",
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        self.chk_delete = ttk.Checkbutton(
            out_frame,
            text="Delete original file",
            variable=self.var_delete_original,
        )
        self.chk_delete.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))

        # Action section
        action_frame = ttk.Frame(self)
        action_frame.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        action_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(action_frame, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=0, sticky="ew")

        self.btn_cancel = ttk.Button(
            action_frame, text="Cancel", command=self._on_cancel, state="disabled"
        )
        self.btn_cancel.grid(row=0, column=1, padx=(12, 0))

        self.btn_process = ttk.Button(action_frame, text="Process", command=self._on_process)
        self.btn_process.grid(row=0, column=2, padx=(8, 0))

        ttk.Label(self, textvariable=self.var_status, wraplength=560, justify="left").grid(
            row=4, column=0, sticky="w", pady=(12, 0)
        )

        log_path = app_logger.log_file_path()
        ttk.Label(
            self,
            text=f"Log: {log_path}" if log_path else "Log tidak tersedia.",
            foreground="#777777",
            wraplength=560,
            justify="left",
        ).grid(row=5, column=0, sticky="w", pady=(12, 0))

    # -- state -------------------------------------------------------------

    def _update_controls(self) -> None:
        has_file = self._source is not None
        analysed = self._total_pages is not None
        busy = self._busy

        self.btn_select.state(["disabled"] if busy else ["!disabled"])
        self.btn_analyze.state(["!disabled"] if has_file and not busy else ["disabled"])
        self.btn_process.state(
            ["!disabled"] if has_file and analysed and not busy else ["disabled"]
        )
        self.btn_cancel.state(["!disabled"] if busy else ["disabled"])
        state = "disabled" if busy else "normal"
        for widget in (self.entry_page_a, self.entry_page_b):
            widget.configure(state=state)
        self.chk_delete.state(["disabled"] if busy else ["!disabled"])

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        if busy:
            self.progress.configure(mode="indeterminate")
            self.progress.start(15)
        else:
            self.progress.stop()
            self.progress.configure(mode="determinate", value=0)
        self._update_controls()

    def _status(self, message: str) -> None:
        self.var_status.set(message)

    # -- events ------------------------------------------------------------

    def _on_select_file(self) -> None:
        if self._busy:
            return
        selected = filedialog.askopenfilename(
            title="Select Word File",
            filetypes=[("Word Document", "*.docx"), ("All files", "*.*")],
        )
        if not selected:
            return

        try:
            path = validation.validate_input_file(selected)
        except ValidationError as exc:
            self._reset_file()
            self.var_file_status.set("Tidak valid.")
            messagebox.showerror("File tidak valid", str(exc), parent=self.master_window)
            return

        self._source = path
        self._total_pages = None
        info = utils.FileInfo(path=path, size_bytes=path.stat().st_size)
        self.var_file_name.set(info.name)
        self.var_file_location.set(str(info.directory))
        self.var_file_size.set(info.size_human)
        self.var_file_status.set("Valid. Siap dianalisis.")
        self.var_pages.set("-")
        LOGGER.info("File dipilih: %s (%s)", path, info.size_human)
        self._update_controls()
        self._on_analyze()

    def _reset_file(self) -> None:
        self._source = None
        self._total_pages = None
        self.var_file_name.set("Belum ada file dipilih.")
        self.var_file_location.set("-")
        self.var_file_size.set("-")
        self.var_pages.set("-")
        self._update_controls()

    def _on_analyze(self) -> None:
        if self._busy or self._source is None:
            return
        source = self._source
        self._cancel_event.clear()
        self._set_busy(True)
        self._status("Analyzing document...")

        def work() -> None:
            try:
                result = split_engine.analyze_document(
                    source,
                    progress=lambda msg: self._queue.put(("progress", msg)),
                    cancel_event=self._cancel_event,
                )
                self._queue.put(("analysis", result))
            except OperationCancelled:
                self._queue.put(("cancelled", None))
            except (SplitError, ValidationError) as exc:
                self._queue.put(("error", ("Analisis gagal", str(exc))))
            except Exception as exc:  # noqa: BLE001
                LOGGER.exception("Kegagalan tak terduga saat analisis.")
                self._queue.put(
                    (
                        "error",
                        (
                            "Analisis gagal",
                            "Terjadi kesalahan tak terduga saat menganalisis dokumen. "
                            "Detail teknis telah dicatat pada file log.\n\n"
                            f"{type(exc).__name__}",
                        ),
                    )
                )

        self._start_worker(work, "AnalyzeWorker")

    def _on_process(self) -> None:
        if self._busy or self._source is None or self._total_pages is None:
            return

        try:
            page_a = validation.parse_page_number(self.var_page_a.get(), "Page A")
            page_b = validation.parse_page_number(self.var_page_b.get(), "Page B")
            plan = validation.compute_split_plan(page_a, page_b, self._total_pages)
            output_one, output_two = validation.resolve_output_paths(
                self._source, self.var_name_one.get(), self.var_name_two.get()
            )
        except ValidationError as exc:
            messagebox.showerror("Input tidak valid", str(exc), parent=self.master_window)
            return

        overwrite = False
        conflicts = validation.existing_outputs(output_one, output_two)
        if conflicts:
            dialog = ConflictDialog(self.master_window, conflicts)
            self.master_window.wait_window(dialog)
            if dialog.choice == "cancel":
                self._status("Proses dibatalkan oleh pengguna.")
                return
            if dialog.choice == "rename":
                self._status("Silakan ubah File 1 Name atau File 2 Name lalu tekan Process.")
                return
            overwrite = True

        if self.var_delete_original.get():
            confirmed = messagebox.askyesno(
                "Konfirmasi penghapusan",
                "Opsi 'Delete original file' aktif. File asli akan dihapus setelah "
                "kedua file hasil berhasil dibuat dan divalidasi.\n\n"
                f"File asli: {self._source.name}\n\nLanjutkan?",
                parent=self.master_window,
            )
            if not confirmed:
                self._status("Proses dibatalkan oleh pengguna.")
                return

        request = SplitRequest(
            source=self._source,
            page_a=page_a,
            page_b=page_b,
            output_one=output_one,
            output_two=output_two,
            delete_original=self.var_delete_original.get(),
            overwrite_allowed=overwrite,
        )
        LOGGER.info(
            "Permintaan pemisahan: source=%s | %s | output=%s, %s | hapus asli=%s",
            request.source, plan.description, output_one.name, output_two.name,
            request.delete_original,
        )

        self._cancel_event.clear()
        self._set_busy(True)
        self._status("Processing...")

        def work() -> None:
            try:
                result = split_engine.perform_split(
                    request,
                    progress=lambda msg: self._queue.put(("progress", msg)),
                    cancel_event=self._cancel_event,
                )
                self._queue.put(("success", result))
            except OperationCancelled:
                self._queue.put(("cancelled", None))
            except (SplitError, ValidationError) as exc:
                self._queue.put(("error", ("Proses gagal", str(exc))))
            except Exception as exc:  # noqa: BLE001
                LOGGER.exception("Kegagalan tak terduga saat proses pemisahan.")
                self._queue.put(
                    (
                        "error",
                        (
                            "Proses gagal",
                            "Terjadi kesalahan tak terduga. File asli tidak diubah. "
                            "Detail teknis telah dicatat pada file log.\n\n"
                            f"{type(exc).__name__}",
                        ),
                    )
                )

        self._start_worker(work, "SplitWorker")

    def _start_worker(self, target, name: str) -> None:
        self._worker = threading.Thread(target=target, name=name, daemon=True)
        self._worker.start()

    def _on_cancel(self) -> None:
        if not self._busy:
            return
        self._cancel_event.set()
        self._status("Membatalkan setelah tahap yang sedang berjalan selesai...")

    # -- queue -------------------------------------------------------------

    def _drain_queue(self) -> None:
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                self._handle_message(kind, payload)
        except queue.Empty:
            pass
        finally:
            self.after(POLL_INTERVAL_MS, self._drain_queue)

    def _handle_message(self, kind: str, payload: Any) -> None:
        if kind == "progress":
            self._status(str(payload))
            return

        if kind == "analysis":
            self._set_busy(False)
            self._total_pages = payload.total_pages
            self.var_pages.set(str(payload.total_pages))
            self.var_file_status.set("Valid. Pagination berhasil dianalisis.")
            if payload.total_pages >= 2:
                if not self.var_page_a.get().strip():
                    self.var_page_a.set("1")
                if not self.var_page_b.get().strip():
                    self.var_page_b.set("2")
            self._status(
                f"Analisis selesai. Dokumen memiliki {payload.total_pages} halaman."
            )
            self._update_controls()
            return

        if kind == "success":
            self._set_busy(False)
            self._show_success(payload)
            return

        if kind == "cancelled":
            self._set_busy(False)
            self._status("Proses dibatalkan. Tidak ada file yang diubah.")
            messagebox.showinfo(
                "Dibatalkan",
                "Proses dibatalkan. File asli tidak diubah dan file sementara telah dibersihkan.",
                parent=self.master_window,
            )
            return

        if kind == "error":
            self._set_busy(False)
            title, message = payload
            self._status("Proses gagal. Lihat pesan kesalahan.")
            log_path = app_logger.log_file_path()
            suffix = f"\n\nLog: {log_path}" if log_path else ""
            messagebox.showerror(title, f"{message}{suffix}", parent=self.master_window)
            return

    def _show_success(self, result: SplitResult) -> None:
        lines = [
            "Split completed successfully.",
            "",
            f"{result.plan.description}",
            f"{result.output_one.name} : {result.pages_one} halaman",
            f"{result.output_two.name} : {result.pages_two} halaman",
            f"Lokasi: {result.output_one.parent}",
        ]
        if result.original_deleted:
            lines.append("File asli telah dihapus.")
        elif not self.var_delete_original.get():
            lines.append("File asli tetap disimpan.")
        if result.warnings:
            lines.append("")
            lines.extend(f"Peringatan: {item}" for item in result.warnings)

        self._status("Split completed successfully.")
        if result.warnings:
            messagebox.showwarning(
                "Selesai dengan peringatan", "\n".join(lines), parent=self.master_window
            )
        else:
            messagebox.showinfo("Selesai", "\n".join(lines), parent=self.master_window)

        if result.original_deleted:
            self._reset_file()

    # -- shutdown ----------------------------------------------------------

    def on_close(self) -> None:
        if self._busy:
            if not messagebox.askyesno(
                "Proses sedang berjalan",
                "Sebuah proses masih berjalan. Menutup aplikasi sekarang akan "
                "membatalkan proses tersebut.\n\nTetap tutup?",
                parent=self.master_window,
            ):
                return
            self._cancel_event.set()
        LOGGER.info("Aplikasi ditutup oleh pengguna.")
        self.master_window.destroy()


def build_root() -> tk.Tk:
    root = tk.Tk()
    root.title(APP_TITLE)
    root.minsize(*WINDOW_MIN_SIZE)
    try:
        style = ttk.Style(root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except tk.TclError:
        pass
    app = WordSplitterApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    return root


def run() -> None:
    root = build_root()
    root.mainloop()

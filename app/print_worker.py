"""
Background print worker threads, job dataclass, and multi-printer coordinator.

Features:
- Single-printer or parallel multi-printer load balancing.
- Automatic failover routing to a backup printer on error.
- Binding margin offset and reverse page order support.
- Pause / Resume / Cancel signals.
- Thread-safe callbacks for UI updates.
"""

import os
import queue
import threading
from dataclasses import dataclass
from typing import Callable, Optional

from app.printer_manager import PrinterManager
from app.settings import FileStatus


@dataclass
class PrintJob:
    """All information needed to print a single PDF file or separator sheet."""
    index: int               # Row index in the GUI file list (-1 for separator sheets)
    pdf_path: str
    filename: str
    pages: list[int]          # 0-indexed page numbers
    copies: int
    paper_size: str           # Key, e.g. "A4"
    orientation: int          # win32con constant
    duplex: int               # win32con constant
    fit_to_page: bool
    binding_margin_mm: float = 0.0
    reverse_order: bool = False
    is_separator: bool = False
    failover_retried: bool = False


class PrintWorker(threading.Thread):
    """
    Worker thread that pulls jobs from a shared queue and prints to a specific printer.
    """

    def __init__(
        self,
        printer_name: str,
        job_queue: queue.Queue,
        cancel_event: threading.Event,
        pause_event: threading.Event,
        coordinator: "MultiPrinterCoordinator",
    ):
        super().__init__(daemon=True)
        self.printer_name = printer_name
        self.job_queue = job_queue
        self.cancel_event = cancel_event
        self.pause_event = pause_event
        self.coordinator = coordinator

    def run(self):
        while not self.cancel_event.is_set():
            try:
                job: PrintJob = self.job_queue.get(timeout=0.5)
            except queue.Empty:
                # Tất cả job đã được enqueue từ đầu; chỉ thoát khi
                # coordinator xác nhận đã xong hết hoặc đã cancel.
                if self.cancel_event.is_set():
                    break
                if self.coordinator.completed_count >= self.coordinator.total_count:
                    break
                # Chờ job failover / job bổ sung, tránh thoát sớm
                continue

            # ── Check Cancel / Pause ─────────────────────────────────
            if self.cancel_event.is_set():
                self.coordinator.notify_file_complete(
                    job.index, job.filename, FileStatus.CANCELLED, self.printer_name
                )
                self.job_queue.task_done()
                continue

            self.pause_event.wait()
            if self.cancel_event.is_set():
                self.coordinator.notify_file_complete(
                    job.index, job.filename, FileStatus.CANCELLED, self.printer_name
                )
                self.job_queue.task_done()
                continue

            # ── Notify Start ─────────────────────────────────────────
            self.coordinator.notify_file_start(job.index, job.filename, self.printer_name)

            # ── Print ────────────────────────────────────────────────
            success = False
            error_msg = ""
            try:
                def _page_cb(cur: int, total: int, _job=job):
                    self.coordinator.notify_page_progress(
                        _job.index,
                        _job.filename,
                        cur,
                        total,
                        self.coordinator.completed_count,
                        self.coordinator.total_count,
                    )

                PrinterManager.print_pdf(
                    printer_name=self.printer_name,
                    pdf_path=job.pdf_path,
                    pages=job.pages,
                    copies=job.copies,
                    paper_size_name=job.paper_size,
                    orientation_value=job.orientation,
                    duplex_value=job.duplex,
                    fit_to_page=job.fit_to_page,
                    cancel_event=self.cancel_event,
                    page_callback=_page_cb,
                    binding_margin_mm=job.binding_margin_mm,
                    reverse_order=job.reverse_order,
                )

                if self.cancel_event.is_set():
                    status = FileStatus.CANCELLED
                else:
                    status = FileStatus.PRINTED
                    success = True

                self.coordinator.increment_completed()
                self.coordinator.notify_file_complete(
                    job.index, job.filename, status, self.printer_name
                )

            except Exception as exc:
                error_msg = str(exc)
                # ── Failover Routing ─────────────────────────────────
                backup_printer = self.coordinator.failover_printer
                if (
                    backup_printer
                    and backup_printer != self.printer_name
                    and not job.failover_retried
                    and not self.cancel_event.is_set()
                ):
                    job.failover_retried = True
                    self.coordinator.notify_failover(
                        job.filename, self.printer_name, backup_printer, error_msg
                    )
                    # Retry printing directly to backup printer
                    try:
                        PrinterManager.print_pdf(
                            printer_name=backup_printer,
                            pdf_path=job.pdf_path,
                            pages=job.pages,
                            copies=job.copies,
                            paper_size_name=job.paper_size,
                            orientation_value=job.orientation,
                            duplex_value=job.duplex,
                            fit_to_page=job.fit_to_page,
                            cancel_event=self.cancel_event,
                            page_callback=_page_cb,
                            binding_margin_mm=job.binding_margin_mm,
                            reverse_order=job.reverse_order,
                        )
                        if self.cancel_event.is_set():
                            status = FileStatus.CANCELLED
                        else:
                            status = FileStatus.PRINTED
                            success = True
                        self.coordinator.increment_completed()
                        self.coordinator.notify_file_complete(
                            job.index, job.filename, status, backup_printer
                        )
                    except Exception as fail_exc:
                        self.coordinator.increment_completed()
                        self.coordinator.notify_file_complete(
                            job.index, job.filename, FileStatus.ERROR, backup_printer
                        )
                        self.coordinator.notify_error(
                            job.index, job.filename, f"Lỗi cả máy in chính & dự phòng: {fail_exc}"
                        )
                else:
                    self.coordinator.increment_completed()
                    self.coordinator.notify_file_complete(
                        job.index, job.filename, FileStatus.ERROR, self.printer_name
                    )
                    self.coordinator.notify_error(
                        job.index, job.filename, error_msg
                    )

            # Cleanup temp separator sheet if applicable
            if job.is_separator and os.path.exists(job.pdf_path):
                try:
                    os.remove(job.pdf_path)
                except Exception:
                    pass

            self.job_queue.task_done()


class MultiPrinterCoordinator:
    """
    Coordinates printing across one or multiple printers in parallel,
    with failover support and unified progress reporting.
    """

    def __init__(
        self,
        jobs: list[PrintJob],
        printer_names: list[str],
        failover_printer: Optional[str] = None,
        on_file_start: Optional[Callable] = None,
        on_page_progress: Optional[Callable] = None,
        on_file_complete: Optional[Callable] = None,
        on_all_complete: Optional[Callable] = None,
        on_error: Optional[Callable] = None,
        on_failover: Optional[Callable] = None,
    ):
        self.jobs = jobs
        self.printer_names = [p for p in printer_names if p]
        self.failover_printer = failover_printer if (failover_printer and failover_printer not in ("", "(Không dùng)")) else None

        # Callbacks
        self.on_file_start = on_file_start          # (index, filename, printer_name)
        self.on_page_progress = on_page_progress    # (index, filename, cur, total, done, ftotal)
        self.on_file_complete = on_file_complete    # (index, filename, status, printer_name)
        self.on_all_complete = on_all_complete      # (was_cancelled: bool)
        self.on_error = on_error                    # (index, filename, error_msg)
        self.on_failover = on_failover              # (filename, from_printer, to_printer, error_msg)

        # Control
        self.cancel_event = threading.Event()
        self.pause_event = threading.Event()
        self.pause_event.set()

        self.job_queue: queue.Queue = queue.Queue()
        for j in jobs:
            self.job_queue.put(j)

        self._completed_lock = threading.Lock()
        self._completed = 0
        self._total = len(jobs)

        self._workers: list[PrintWorker] = []
        self._monitor_thread: Optional[threading.Thread] = None

    @property
    def completed_count(self) -> int:
        with self._completed_lock:
            return self._completed

    @property
    def total_count(self) -> int:
        return self._total

    def increment_completed(self):
        with self._completed_lock:
            self._completed += 1

    def notify_file_start(self, index: int, filename: str, printer_name: str):
        if self.on_file_start:
            self.on_file_start(index, filename, printer_name)

    def notify_page_progress(self, index: int, filename: str, cur: int, total: int, done: int, ftotal: int):
        if self.on_page_progress:
            self.on_page_progress(index, filename, cur, total, done, ftotal)

    def notify_file_complete(self, index: int, filename: str, status: str, printer_name: str):
        if self.on_file_complete:
            self.on_file_complete(index, filename, status, printer_name)

    def notify_error(self, index: int, filename: str, error_msg: str):
        if self.on_error:
            self.on_error(index, filename, error_msg)

    def notify_failover(self, filename: str, from_printer: str, to_printer: str, error_msg: str):
        if self.on_failover:
            self.on_failover(filename, from_printer, to_printer, error_msg)

    def start(self):
        """Start worker threads and background completion monitor."""
        if not self.printer_names:
            raise ValueError("Không có máy in nào được chọn.")

        for pname in self.printer_names:
            worker = PrintWorker(
                printer_name=pname,
                job_queue=self.job_queue,
                cancel_event=self.cancel_event,
                pause_event=self.pause_event,
                coordinator=self,
            )
            self._workers.append(worker)
            worker.start()

        self._monitor_thread = threading.Thread(target=self._monitor_completion, daemon=True)
        self._monitor_thread.start()

    def _monitor_completion(self):
        for w in self._workers:
            w.join()
        if self.on_all_complete:
            self.on_all_complete(self.cancel_event.is_set())

    def cancel(self):
        self.cancel_event.set()
        self.pause_event.set()

    def pause(self):
        self.pause_event.clear()

    def resume(self):
        self.pause_event.set()

    @property
    def is_paused(self) -> bool:
        return not self.pause_event.is_set()

    @property
    def is_cancelled(self) -> bool:
        return self.cancel_event.is_set()


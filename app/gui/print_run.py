from __future__ import annotations

import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from typing import Optional

import customtkinter as ctk
import win32con

from app.settings import (
    APP_NAME, WINDOW_TITLE, DEFAULT_SIZE, MIN_SIZE,
    PAPER_SIZES, CUSTOM_PAPER_LABEL, DUPLEX_MODES, ORIENTATIONS, THEME_COLORS,
    PAGE_RANGE_ALL, PAGE_RANGE_CUSTOM, PAGE_RANGE_ODD, PAGE_RANGE_EVEN,
    PAGE_RANGE_OPTIONS, PAGE_MODE_BY_LABEL, PAGE_LABEL_BY_MODE,
    ORIENT_AUTO, ORIENT_PORTRAIT, ORIENT_LANDSCAPE, ORIENTATION_OPTIONS,
    INVOICE_OFF, INVOICE_AUTO, INVOICE_WARN, INVOICE_MODES,
    BINDING_MARGIN_OPTIONS,
    SUPPORTED_EXTENSIONS, FILE_DIALOG_TYPES,
    FileStatus, STATUS_ICONS, STATUS_COLORS,
    parse_custom_paper, resolve_paper,
)
from app.utils import parse_page_range, format_file_size, get_timestamp, resolve_page_selection
from app.pdf_manager import PDFManager, FileInfo
from app.printer_manager import PrinterManager
from app.print_worker import PrintJob, PrintWorker, MultiPrinterCoordinator
from app.preview import PreviewWindow
from app.file_converter import cleanup_temp
from app.config_store import CONFIG_FILE, get_app_data_dir, load_config, save_config
from app.queue_store import apply_move, dedupe_paths, is_edit_locked, queue_summary, retry_indices
from app.app_logger import log_info, log_error, get_log_dir
from app import updater as app_updater
from app import invoice_detect as inv_detect
from PIL import Image, ImageTk
import sys


"""Part of PDFBatchPrinterApp (see app/gui/__init__.py)."""


class PrintRunMixin:

    def _finish_start_print(self, jobs: list[PrintJob], active_indices: list[int],
                            printers_to_use: list[str], failover_p: Optional[str],
                            duplex_name: str, paper: str, orient_name: str,
                            reverse_order: bool, remove_blanks: bool,
                            use_separator: bool, total_pages: int,
                            total_copies_pages: int, total_blank_skipped: int,
                            action_title: str, auto_duplex_files: Optional[list] = None,
                            saved_sheets: int = 0):

        # Confirmation Dialog
        p_desc = f"{len(printers_to_use)} máy in ({', '.join(printers_to_use)})" if len(printers_to_use) > 1 else printers_to_use[0]
        f_desc = failover_p if failover_p else "Không kích hoạt"
        blank_desc = f"Bật (đã loại bỏ {total_blank_skipped} trang trắng)" if remove_blanks else "Tắt"
        rev_desc = "Trang cuối về trang 1" if reverse_order else "Trang 1 đến trang cuối"
        sep_desc = "Có chèn tờ bìa phân cách" if use_separator else "Không chèn"
        auto_duplex_files = auto_duplex_files or []
        if auto_duplex_files:
            inv_desc = (f"{len(auto_duplex_files)} hóa đơn tự 2 mặt "
                        f"(tiết kiệm ~{saved_sheets} tờ)")
        else:
            inv_desc = "Không có"

        msg = (
            f"• Số tệp tin:             {len(active_indices)}\n"
            f"• Tổng số trang in:       {total_pages}\n"
            f"• Tổng trang × bản:        {total_copies_pages}\n"
            f"• Máy in thực hiện:       {p_desc}\n"
            f"• Máy in dự phòng:        {f_desc}\n"
            f"• Kiểu in 2 mặt:          {duplex_name}\n"
            f"• Khổ giấy / Chiều:       {paper} / {orient_name}\n"
            f"• Lề đóng gáy sách:       {self._binding_margin_var.get()}\n"
            f"• Thứ tự trang:           {rev_desc}\n"
            f"• Bỏ trang trắng:         {blank_desc}\n"
            f"• Trang bìa phân cách:    {sep_desc}\n"
            f"• 🧾 Hóa đơn 2 mặt:       {inv_desc}\n"
        )
        if not messagebox.askyesno(action_title, msg, icon="question"):
            return

        # Auto save settings
        self._save_config()

        # Reset status for active target files
        for i in active_indices:
            if 0 <= i < len(self.file_list):
                self.file_list[i].status = FileStatus.WAITING
        self._refresh_tree(preserve_selection=True)
        self._reset_progress()
        self._set_printing_state(True)

        self._log(f"Bắt đầu in {len(active_indices)} tệp trên [{p_desc}]")
        self._log(f"Thiết lập: {duplex_name} | {paper} | {orient_name} | Gáy: {self._binding_margin_var.get()}")

        self.print_worker = MultiPrinterCoordinator(
            jobs=jobs,
            printer_names=printers_to_use,
            failover_printer=failover_p,
            on_file_start=self._cb_file_start,
            on_page_progress=self._cb_page_progress,
            on_file_complete=self._cb_file_complete,
            on_all_complete=self._cb_all_complete,
            on_error=self._cb_error,
            on_failover=self._cb_failover,
        )
        self.print_worker.start()

    def start_print_selected(self):
        """Prints only the currently selected items in the queue."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(
                "Chỉ in mục chọn",
                "Vui lòng chọn các tệp tin trong danh sách bạn muốn in.",
            )
            return

        indices = []
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    indices.append(idx)
            except Exception:
                pass

        if not indices:
            return

        if any(self.file_list[i].status == FileStatus.PRINTING for i in indices):
            messagebox.showwarning("Cảnh báo", "Có tệp tin trong danh sách chọn đang được in.")
            return

        self.start_print(target_indices=indices, action_title="XÁC NHẬN IN CÁC MỤC ĐÃ CHỌN")

    def retry_failed_prints(self):
        """Finds and reprints all files that failed or were cancelled."""
        indices = retry_indices(self.file_list)
        if not indices:
            messagebox.showinfo(
                "In lại tệp lỗi",
                "Không có tệp tin nào bị lỗi hoặc bị hủy trong hàng đợi.",
            )
            return

        self.start_print(target_indices=indices, action_title="XÁC NHẬN IN LẠI CÁC TỆP LỖI / ĐÃ HỦY")

    def _reset_selected_to_waiting(self):
        """Resets selected items' status back to Waiting (Chờ in)."""
        sel = self.tree.selection()
        if not sel:
            return
        count = 0
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    if self.file_list[idx].status != FileStatus.PRINTING:
                        self.file_list[idx].status = FileStatus.WAITING
                        count += 1
            except Exception:
                pass
        if count:
            self._refresh_tree(preserve_selection=True)
            self._log(f"Đã đặt lại trạng thái 'Chờ in' cho {count} tệp tin")

    def pause_print(self):
        if self.print_worker is None:
            return
        if self.print_worker.is_paused:
            self.print_worker.resume()
            self.btn_pause.configure(text="⏸️ Tạm Dừng")
            self.status_lbl.configure(text="⏳ Đang in...")
            self._log("Tiếp tục in")
        else:
            self.print_worker.pause()
            self.btn_pause.configure(text="▶️ Tiếp Tục")
            self.status_lbl.configure(text="⏸️ Đang tạm dừng...")
            self._log("Tạm dừng in")

    def cancel_print(self):
        if self.print_worker is None:
            return
        if not messagebox.askyesno("Hủy in", "Bạn có chắc muốn dừng và hủy toàn bộ quá trình in?"):
            return
        self.print_worker.cancel()
        self._log("Đã gửi lệnh hủy in")

    # ── Worker Callbacks (Thread-Safe via Queue) ─────────────────────

    def _cb_file_start(self, idx, fn, printer_name=""):
        self._async_queue.put((self._on_file_start, (idx, fn, printer_name)))

    def _cb_page_progress(self, idx, fn, cur, tot, done, ftot):
        self._async_queue.put((self._on_page_progress, (idx, fn, cur, tot, done, ftot)))

    def _cb_file_complete(self, idx, fn, st, printer_name=""):
        self._async_queue.put((self._on_file_complete, (idx, fn, st, printer_name)))

    def _cb_all_complete(self, cancelled):
        self._async_queue.put((self._on_all_complete, (cancelled,)))

    def _cb_error(self, idx, fn, msg):
        self._async_queue.put((self._on_error, (idx, fn, msg)))

    def _cb_failover(self, fn, from_p, to_p, err):
        self._async_queue.put((self._on_failover, (fn, from_p, to_p, err)))

    # ── GUI Handlers ─────────────────────────────────────────────────

    def _on_file_start(self, index, filename, printer_name=""):
        p_tag = f" [{printer_name}]" if printer_name else ""
        if 0 <= index < len(self.file_list):
            self.file_list[index].status = FileStatus.PRINTING
            self._refresh_tree(preserve_selection=True)
            self.status_lbl.configure(text=f"⏳ Đang in: {filename}{p_tag}")
            self._log(f"{filename} — bắt đầu in ({self.file_list[index].copies} bản){p_tag}")
        elif index == -1:
            self.status_lbl.configure(text=f"📄 Đang in tờ bìa phân cách{p_tag}")
            self._log(f"📄 Bìa phân cách — bắt đầu in{p_tag}")
        self.file_progress.set(0)
        self.file_progress_lbl.configure(text="0 trang")

    def _on_page_progress(self, index, filename, cur, total, done, ftotal):
        self.file_progress.set(cur / total if total else 0)
        self.file_progress_lbl.configure(text=f"{cur}/{total} trang")
        prog = (done + cur / total) / ftotal if ftotal else 0
        self.total_progress.set(prog)
        self.total_progress_lbl.configure(text=f"{done}/{ftotal} mục")

    def _on_file_complete(self, index, filename, status, printer_name=""):
        p_tag = f" [{printer_name}]" if printer_name else ""
        if 0 <= index < len(self.file_list):
            self.file_list[index].status = status
            self._refresh_tree(preserve_selection=True)
        icon = STATUS_ICONS.get(status, "")
        self._log(f"{filename} — {icon} {status}{p_tag}")

    def _on_failover(self, filename, from_p, to_p, err):
        self._log(f"🛡️ TỰ ĐỘNG CHUYỂN DỰ PHÒNG: [{filename}] gặp sự cố trên '{from_p}' ({err}) → Đã tự động chuyển sang máy in '{to_p}'!")
        self.status_lbl.configure(text=f"🛡️ Chuyển dự phòng: {filename} → {to_p}")

    def _on_all_complete(self, was_cancelled):
        self._set_printing_state(False)
        if was_cancelled:
            for f in self.file_list:
                if f.status == FileStatus.WAITING:
                    f.status = FileStatus.CANCELLED
            self._refresh_tree(preserve_selection=True)
            self.status_lbl.configure(text="⏹️ Đã hủy quá trình in")
            self._log("Quá trình in đã bị hủy")
        else:
            self.total_progress.set(1.0)
            self.status_lbl.configure(text="✓ Hoàn thành tất cả tệp tin!")
            self._log("✓ Hoàn thành in tất cả tệp tin")

    def _on_error(self, index, filename, error_msg):
        self._log(f"⚠ LỖI [{filename}]: {error_msg}")

    # ── Progress Helpers ─────────────────────────────────────────────

    def _reset_progress(self):
        self.file_progress.set(0)
        self.file_progress_lbl.configure(text="0/0 trang")
        self.total_progress.set(0)
        self.total_progress_lbl.configure(text="0/0 tệp")
        self.status_lbl.configure(text="⏳ Đang khởi tạo...")

    def _set_printing_state(self, printing: bool):
        self._is_printing = printing
        st = "disabled" if printing else "normal"
        self.btn_start.configure(state=st)
        self.btn_print_selected.configure(state=st)
        self.btn_retry_failed.configure(state=st)
        self.btn_pause.configure(state="normal" if printing else "disabled")
        self.btn_cancel.configure(state="normal" if printing else "disabled")
        if printing:
            self.btn_pause.configure(text="⏸️ Tạm Dừng")
        else:
            self.print_worker = None

    # ── Logging ──────────────────────────────────────────────────────

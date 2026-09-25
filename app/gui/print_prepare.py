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


class PrintPrepareMixin:

    def start_print(self, target_indices: Optional[list[int]] = None, action_title: str = "XÁC NHẬN BẮT ĐẦU IN"):
        if not self.file_list:
            messagebox.showwarning("Cảnh báo", "Hàng đợi trống — chưa có file nào.")
            return

        if self._is_printing:
            messagebox.showwarning("Cảnh báo", "Quá trình in đang diễn ra. Vui lòng chờ hoàn thành hoặc bấm Hủy in trước.")
            return

        # Determine target printers
        primary_printer = self._printer_var.get()
        if self._parallel_printers and len(self._parallel_printers) >= 2:
            printers_to_use = self._parallel_printers
        elif primary_printer and not primary_printer.startswith("("):
            printers_to_use = [primary_printer]
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn máy in trước khi bắt đầu.")
            return

        failover_p = self._failover_printer_var.get()
        if failover_p in ("", "(Không dùng)"):
            failover_p = None

        orient_name = self._orient_var.get()
        duplex_name = self._duplex_var.get()
        duplex_val = DUPLEX_MODES.get(duplex_name, win32con.DMDUP_SIMPLEX)
        paper = self._resolve_gui_paper()
        fit = self._fit_to_page.get()
        binding_margin_mm = float(BINDING_MARGIN_OPTIONS.get(self._binding_margin_var.get(), 0))
        reverse_order = self._reverse_order_var.get()
        remove_blanks = self._remove_blanks_var.get()
        use_separator = self._separator_sheet_var.get()

        active_indices = target_indices if target_indices is not None else list(range(len(self.file_list)))
        if not active_indices:
            messagebox.showinfo("Thông báo", "Không có tệp tin nào để in.")
            return

        # Chạy chuẩn bị file (ensure_pdf/convert Office) dưới nền để không treo UI.
        prep_win = ctk.CTkToplevel(self)
        prep_win.title("Đang chuẩn bị tệp in...")
        prep_win.geometry("380x130")
        prep_win.transient(self)
        prep_win.grab_set()
        prep_win.resizable(False, False)
        ctk.CTkLabel(
            prep_win, text="⏳ Đang chuẩn bị tệp in (chuyển đổi Office → PDF)...",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        ).pack(padx=16, pady=(18, 8))
        prep_lbl = ctk.CTkLabel(prep_win, text=f"0/{len(active_indices)}", font=ctk.CTkFont(size=11))
        prep_lbl.pack(padx=16, pady=(0, 6))
        prep_bar = ctk.CTkProgressBar(prep_win, width=320)
        prep_bar.pack(padx=16, pady=(0, 14))
        prep_bar.set(0)
        self.configure(cursor="wait")

        # Snapshot cài đặt để thread nền không chạm widget Tk.
        # Trang in là thuộc tính RIÊNG từng file (finfo.page_mode/value).
        # Duplex hiệu dụng: override người dùng > auto hóa đơn > mặc định chung.
        _printers_snapshot = list(printers_to_use)
        _invoice_mode = self._invoice_mode_var.get()
        _duplex_file_map = {
            "simplex": win32con.DMDUP_SIMPLEX,
            "long": win32con.DMDUP_VERTICAL,
            "short": win32con.DMDUP_HORIZONTAL,
        }

        import threading

        def _bg_prepare():
            jobs: list[PrintJob] = []
            total_pages = 0
            total_copies_pages = 0
            total_blank_skipped = 0
            blank_logs: list[str] = []
            warn_suggest: list[int] = []  # file idx hóa đơn chờ hỏi (chế độ cảnh báo)
            auto_duplex_files: list[str] = []
            saved_sheets = 0
            error: Optional[str] = None

            for doc_num, i in enumerate(active_indices, 1):
                if i < 0 or i >= len(self.file_list):
                    continue
                finfo = self.file_list[i]
                try:
                    self.pdf_manager.ensure_pdf(finfo)
                except Exception as conv_exc:
                    error = f"Không thể chuẩn bị file: {finfo.filename}\n{conv_exc}"
                    break

                # Nhận diện bù cho file Office/scan (PDF đã làm lúc nạp)
                if _invoice_mode != INVOICE_OFF and not finfo.invoice_analyzed:
                    try:
                        self._analyze_invoice(finfo)
                    except Exception:
                        pass

                # 1. Trang in RIÊNG từng file (không dùng biến chung)
                try:
                    p0 = resolve_page_selection(finfo.page_mode, finfo.page_range_text, finfo.page_count)
                except ValueError as exc:
                    error = f"File: {finfo.filename}\n⚠ {exc}"
                    break

                if remove_blanks:
                    p0, skipped = PDFManager.filter_pages(finfo.pdf_path, p0, remove_blanks=True)
                    if skipped:
                        total_blank_skipped += len(skipped)
                        blank_logs.append(f"[{finfo.filename}] Đã tự động bỏ qua {len(skipped)} trang trắng")

                if not p0:
                    p0 = [0]

                # 2. Duplex hiệu dụng từng file (§6, §11-12, §18)
                file_duplex = duplex_val
                if finfo.duplex_override:
                    if finfo.duplex_mode in _duplex_file_map:
                        file_duplex = _duplex_file_map[finfo.duplex_mode]
                elif finfo.duplex_mode in _duplex_file_map:
                    file_duplex = _duplex_file_map[finfo.duplex_mode]
                    if finfo.duplex_auto:
                        auto_duplex_files.append(finfo.filename)
                        saved_sheets += (len(p0) - (len(p0) + 1) // 2) * finfo.copies
                elif (_invoice_mode == INVOICE_AUTO and finfo.invoice_detected
                        and len(p0) >= 2):
                    file_duplex = win32con.DMDUP_VERTICAL
                    finfo.duplex_mode = "long"
                    finfo.duplex_auto = True
                    auto_duplex_files.append(finfo.filename)
                    saved_sheets += (len(p0) - (len(p0) + 1) // 2) * finfo.copies
                elif (_invoice_mode == INVOICE_WARN and finfo.invoice_detected
                        and len(p0) >= 2):
                    warn_suggest.append(i)

                if orient_name == ORIENT_AUTO:
                    ori = (
                        win32con.DMORIENT_LANDSCAPE
                        if self.pdf_manager.is_landscape(finfo.pdf_path)
                        else win32con.DMORIENT_PORTRAIT
                    )
                else:
                    ori = ORIENTATIONS.get(orient_name, win32con.DMORIENT_PORTRAIT)

                if use_separator and len(active_indices) > 1:
                    try:
                        sep_pdf = PDFManager.create_separator_sheet(
                            filename=finfo.filename,
                            page_count=len(p0),
                            copies=finfo.copies,
                            printer_name=", ".join(_printers_snapshot),
                            doc_index=doc_num,
                        )
                        jobs.append(PrintJob(
                            index=-1, pdf_path=sep_pdf,
                            filename=f"Bìa phân cách #{doc_num}",
                            pages=[0], copies=1, paper_size=paper,
                            orientation=ori, duplex=win32con.DMDUP_SIMPLEX,
                            fit_to_page=True, binding_margin_mm=0.0,
                            reverse_order=False, is_separator=True,
                        ))
                    except Exception as sep_exc:
                        print(f"Error creating separator sheet: {sep_exc}")

                jobs.append(PrintJob(
                    index=i, pdf_path=finfo.pdf_path, filename=finfo.filename,
                    pages=p0, copies=finfo.copies, paper_size=paper,
                    orientation=ori, duplex=file_duplex, fit_to_page=fit,
                    binding_margin_mm=binding_margin_mm,
                    reverse_order=reverse_order, is_separator=False,
                ))
                total_pages += len(p0)
                total_copies_pages += len(p0) * finfo.copies
                self._async_queue.put((lambda d=doc_num, t=len(active_indices): (
                    prep_lbl.configure(text=f"{d}/{t}"),
                    prep_bar.set(d / t if t else 1),
                ), ()))

            def _on_prepared():
                nonlocal saved_sheets
                try:
                    prep_win.destroy()
                except Exception:
                    pass
                self.configure(cursor="")
                if error:
                    messagebox.showerror("Lỗi chuẩn bị tệp", error)
                    return
                for line in blank_logs:
                    self._log(line)
                # Chế độ "Chỉ cảnh báo": hỏi 1 lần duy nhất cho cả đợt (§7, §16)
                if warn_suggest and _invoice_mode == INVOICE_WARN:
                    names = [self.file_list[i].filename for i in warn_suggest
                             if 0 <= i < len(self.file_list)]
                    if names and messagebox.askyesno(
                        "Phát hiện hóa đơn điện tử",
                        f"💡 Phát hiện {len(names)} hóa đơn từ 2 trang thực tế trở lên:\n"
                        + "\n".join(f"• {n}" for n in names[:8])
                        + ("\n..." if len(names) > 8 else "")
                        + "\n\nHệ thống đề xuất: In 2 mặt – Cạnh dài.\nÁp dụng?",
                    ):
                        for i in warn_suggest:
                            if 0 <= i < len(self.file_list):
                                f = self.file_list[i]
                                f.duplex_mode = "long"
                                f.duplex_auto = True
                                auto_duplex_files.append(f.filename)
                        for job in jobs:
                            if job.index in warn_suggest and not job.is_separator:
                                job.duplex = win32con.DMDUP_VERTICAL
                        # Tính lại giấy tiết kiệm cho các file vừa áp dụng
                        for i in warn_suggest:
                            if 0 <= i < len(self.file_list):
                                f = self.file_list[i]
                                try:
                                    n = len(f.resolve_pages())
                                    saved_sheets += (n - (n + 1) // 2) * f.copies
                                except Exception:
                                    pass
                        self._refresh_tree(preserve_selection=True)
                    else:
                        for i in warn_suggest:
                            if 0 <= i < len(self.file_list):
                                self.file_list[i].duplex_override = True
                for name in auto_duplex_files:
                    self._log(f"🧾 {name} — tự động In 2 mặt (Cạnh dài)")
                self._finish_start_print(
                    jobs, active_indices, printers_to_use, failover_p,
                    duplex_name, paper, orient_name, reverse_order,
                    remove_blanks, use_separator,
                    total_pages, total_copies_pages, total_blank_skipped,
                    action_title, auto_duplex_files, saved_sheets,
                )

            self._async_queue.put((_on_prepared, ()))

        threading.Thread(target=_bg_prepare, daemon=True).start()
        return

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


class QueueAddMixin:

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Chọn tệp tin cần in",
            filetypes=FILE_DIALOG_TYPES,
        )
        if paths:
            self._add_to_queue(list(paths))

    def add_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục")
        if not folder:
            return

        if self._include_subfolders.get():
            found = self.pdf_manager.scan_folder(folder, recursive=True)
            if not found:
                messagebox.showinfo("Thông báo", "Không tìm thấy file hỗ trợ trong thư mục.")
                return
            self._add_to_queue(found)
        else:
            paths = filedialog.askopenfilenames(
                title=f"Chọn tệp tin từ: {os.path.basename(folder)}",
                initialdir=folder,
                filetypes=FILE_DIALOG_TYPES,
            )
            if paths:
                self._add_to_queue(list(paths))

    def _add_to_queue(self, paths: list[str]):
        if not paths:
            return
        default_copies = max(1, self._copies_var.get())

        candidates = []
        for p in paths:
            ext = os.path.splitext(p)[1].lower()
            if ext in SUPPORTED_EXTENSIONS and os.path.exists(p):
                candidates.append(p)
        seen = {os.path.abspath(f.original_path) for f in self.file_list}
        valid_paths = dedupe_paths(candidates, seen)

        if not valid_paths:
            self._log("⚠️ Không tìm thấy tệp tin được hỗ trợ.")
            return

        total_to_add = len(valid_paths)
        self.status_lbl.configure(text=f"⚡ Đang nạp {total_to_add} tệp tin...")

        import concurrent.futures
        import threading

        # Snapshot Tk vars ở main thread (thread nền không được chạm Tk)
        _inv_mode = self._invoice_mode_var.get()

        def _bg_loader():
            added_infos = []
            errors = []

            # Parallel metadata parsing
            max_w = min(32, max(4, (os.cpu_count() or 4) * 4))
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_w) as executor:
                futures = [executor.submit(self.pdf_manager.add_file, p, True) for p in valid_paths]
                for p, f in zip(valid_paths, futures):
                    try:
                        info = f.result()
                        info.copies = default_copies
                        added_infos.append(info)
                    except Exception as exc:
                        errors.append(f"{os.path.basename(p)}: {exc}")

            # Nhận diện hóa đơn từng file PDF (tuần tự, nhẹ — text đã có sẵn).
            # _inv_mode đã snapshot ở main thread (Tk var cấm chạm từ nền).
            detected_names = []
            if _inv_mode != INVOICE_OFF:
                for info in added_infos:
                    try:
                        if self._analyze_invoice(info) and info.invoice_detected:
                            detected_names.append(info.filename)
                    except Exception:
                        pass

            def _on_done():
                if not getattr(self, "_is_alive", False):
                    return
                for info in added_infos:
                    self.file_list.append(info)
                    tag = " 🧾 Hóa đơn" if info.invoice_detected else ""
                    self._log(f"+ {info.filename}  ({info.file_type}, {info.page_count} trang){tag}")
                for err in errors:
                    self._log(f"⚠ Lỗi: {err}")

                self._refresh_tree(preserve_selection=True)
                self._update_stats_pill()
                self.status_lbl.configure(text=f"✓ Đã nạp xong {len(added_infos)} tệp tin")

                if added_infos:
                    self._log(f"✓ Đã thêm {len(added_infos)} tệp vào hàng đợi (Tổng: {len(self.file_list)} tệp)")
                    if len(self.file_list) == len(added_infos):
                        children = self.tree.get_children()
                        if children:
                            self.tree.selection_set(children[0])
                            self._on_file_select(None)
                if errors:
                    messagebox.showwarning(
                        "Cảnh báo",
                        f"Không thể thêm {len(errors)} tệp tin:\n\n" + "\n".join(errors[:10]),
                    )

            self._async_queue.put((_on_done, ()))

        threading.Thread(target=_bg_loader, daemon=True).start()

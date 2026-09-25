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


class QueueManageMixin:

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        indices = sorted(
            (int(self.tree.item(iid, "values")[0]) - 1 for iid in sel),
            reverse=True,
        )
        for i in indices:
            if 0 <= i < len(self.file_list):
                del self.file_list[i]
        self._refresh_tree()
        self._update_stats_pill()
        self._update_queue_summary()
        if hasattr(self, "selected_file_lbl"):
            self.selected_file_lbl.configure(text="📄 (Chưa chọn tệp)", text_color=THEME_COLORS["text_muted"])
        self._log(f"Đã xóa {len(indices)} mục khỏi hàng đợi")

    def remove_all(self):
        if not self.file_list:
            return
        n = len(self.file_list)
        self.file_list.clear()
        self._refresh_tree()
        self._update_stats_pill()
        self._update_queue_summary()
        if hasattr(self, "selected_file_lbl"):
            self.selected_file_lbl.configure(text="📄 (Chưa chọn tệp)", text_color=THEME_COLORS["text_muted"])
        self._log(f"Đã xóa tất cả ({n} mục)")

    def _refresh_tree(self, preserve_selection: bool = False):
        selected_indices = []
        if preserve_selection:
            for item in self.tree.selection():
                try:
                    selected_indices.append(int(self.tree.item(item, "values")[0]) - 1)
                except Exception:
                    pass

        self.tree.delete(*self.tree.get_children())
        items_to_select = []
        for i, f in enumerate(self.file_list):
            icon = STATUS_ICONS.get(f.status, "")
            item_id = self.tree.insert("", "end", values=(
                i + 1,
                f.filename,
                f.file_type,
                f.page_count,
                format_file_size(f.file_size),
                f.pages_display(),
                f.copies,
                f"{icon} {f.status}",
                f.doctype_display(),
                f.duplex_display(),
            ))
            if preserve_selection and i in selected_indices:
                items_to_select.append(item_id)

        if items_to_select:
            self.tree.selection_set(items_to_select)

        self._update_queue_summary()

    def _update_stats_pill(self):
        total_files, _copies, total_pages = queue_summary(self.file_list)
        page_text = f"{total_pages} trang in" if total_pages is not None else "Đang tính..."
        self.stats_lbl.configure(text=f"📦 {total_files} tệp tin  •  📄 {page_text}")
        self._update_queue_summary()

    def _update_queue_summary(self):
        total_files, total_copies, total_pages = queue_summary(self.file_list)
        page_text = f"{total_pages} trang" if total_pages is not None else "Đang tính..."

        if hasattr(self, "queue_summary_lbl"):
            self.queue_summary_lbl.configure(
                text=f"📦 Tổng tài liệu: {total_files}  •  📑 Tổng số bản: {total_copies}  •  📄 Tổng số trang in: {page_text}"
            )

    def _on_file_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            if hasattr(self, "selected_file_lbl"):
                self.selected_file_lbl.configure(text="📄 (Chưa chọn tệp)", text_color=THEME_COLORS["text_muted"])
            return
        try:
            idx = int(self.tree.item(sel[0], "values")[0]) - 1
            if 0 <= idx < len(self.file_list):
                finfo = self.file_list[idx]
                if hasattr(self, "selected_file_lbl"):
                    name = finfo.filename if len(finfo.filename) <= 32 else finfo.filename[:31] + "…"
                    self.selected_file_lbl.configure(text=f"📄 {name}", text_color=THEME_COLORS["text"])
                self._copies_var.set(finfo.copies)
                self._sync_page_ui_from_selection()
        except Exception:
            pass

    def move_selected_up(self):
        """Move selected items up one position in the queue."""
        if self._is_printing:
            messagebox.showwarning("Cảnh báo", "Không thể thay đổi thứ tự hàng đợi khi đang in.")
            return

        sel = self.tree.selection()
        if not sel:
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

        new_indices = apply_move(self.file_list, indices, -1)

        self.tree.delete(*self.tree.get_children())
        items_to_select = []
        for i, f in enumerate(self.file_list):
            icon = STATUS_ICONS.get(f.status, "")
            item_id = self.tree.insert("", "end", values=(
                i + 1,
                f.filename,
                f.file_type,
                f.page_count,
                format_file_size(f.file_size),
                f.pages_display(),
                f.copies,
                f"{icon} {f.status}",
                f.doctype_display(),
                f.duplex_display(),
            ))
            if i in new_indices:
                items_to_select.append(item_id)

        if items_to_select:
            self.tree.selection_set(items_to_select)
            self.tree.see(items_to_select[0])

    def move_selected_down(self):
        """Move selected items down one position in the queue."""
        if self._is_printing:
            messagebox.showwarning("Cảnh báo", "Không thể thay đổi thứ tự hàng đợi khi đang in.")
            return

        sel = self.tree.selection()
        if not sel:
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

        new_indices = apply_move(self.file_list, indices, +1)

        self.tree.delete(*self.tree.get_children())
        items_to_select = []
        for i, f in enumerate(self.file_list):
            icon = STATUS_ICONS.get(f.status, "")
            item_id = self.tree.insert("", "end", values=(
                i + 1,
                f.filename,
                f.file_type,
                f.page_count,
                format_file_size(f.file_size),
                f.pages_display(),
                f.copies,
                f"{icon} {f.status}",
                f.doctype_display(),
                f.duplex_display(),
            ))
            if i in new_indices:
                items_to_select.append(item_id)

        if items_to_select:
            self.tree.selection_set(items_to_select)
            self.tree.see(items_to_select[-1])

    def _sort_by_column(self, col_id: str):
        """Sort file queue by clicked column heading with ascending/descending toggle."""
        if self._is_printing:
            messagebox.showwarning("Cảnh báo", "Không thể sắp xếp hàng đợi khi đang in.")
            return

        if not self.file_list:
            return

        is_reverse = self._sort_reverse.get(col_id, False)
        new_reverse = not is_reverse
        self._sort_reverse[col_id] = new_reverse

        key_funcs = {
            "stt": lambda f: 0,
            "filename": lambda f: f.filename.lower(),
            "filetype": lambda f: f.file_type.lower(),
            "pages": lambda f: int(f.page_count) if isinstance(f.page_count, int) else 0,
            "size": lambda f: f.file_size,
            "pagesel": lambda f: f.pages_display(),
            "copies": lambda f: f.copies,
            "status": lambda f: str(f.status),
            "doctype": lambda f: f.document_type,
            "duplex": lambda f: f.duplex_display(),
        }

        k_fn = key_funcs.get(col_id, lambda f: f.filename.lower())
        self.file_list.sort(key=k_fn, reverse=new_reverse)

        headings_map = {
            "stt": "#",
            "filename": "Tên Tệp Tin",
            "filetype": "Định Dạng",
            "pages": "Số Trang",
            "size": "Kích Thước",
            "pagesel": "Trang In",
            "copies": "Số Bản (✎)",
            "status": "Trạng Thái",
            "doctype": "Loại",
            "duplex": "2 Mặt",
        }
        for cid, label in headings_map.items():
            if cid == col_id:
                indicator = " ▼" if new_reverse else " ▲"
                self.tree.heading(cid, text=f"{label}{indicator}")
            else:
                self.tree.heading(cid, text=f"{label} ⇕")

        self._refresh_tree(preserve_selection=True)
        dir_text = "Giảm dần (Z-A)" if new_reverse else "Tăng dần (A-Z)"
        self._log(f"Đã sắp xếp danh sách theo [{headings_map.get(col_id, col_id)}] — {dir_text}")

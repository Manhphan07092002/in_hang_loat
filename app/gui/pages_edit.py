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


class PagesEditMixin:

    def _sync_page_ui_from_selection(self):
        """Show selected file's page setting in the settings panel."""
        if getattr(self, "_syncing_page_ui", False):
            return
        self._syncing_page_ui = True
        try:
            indices = self._selected_indices()
            if not indices:
                if hasattr(self, "page_file_lbl"):
                    self.page_file_lbl.configure(text="📄 Chưa chọn tệp tin")
                self._sync_doc_ui(None)
                return
            first = self.file_list[indices[0]]
            if len(indices) == 1:
                label = f"📄 {first.filename}"
            else:
                label = f"📄 {len(indices)} tệp đã chọn (hiển thị file đầu: {first.filename})"
            if hasattr(self, "page_file_lbl"):
                self.page_file_lbl.configure(text=label if len(label) <= 52 else label[:51] + "…")
            self._page_range_var.set(PAGE_LABEL_BY_MODE.get(first.page_mode, PAGE_RANGE_ALL))
            self._custom_pages_var.set(first.page_range_text)
            self._refresh_custom_pages_visibility()
            self._sync_doc_ui(first)
        finally:
            self._syncing_page_ui = False

    def _refresh_custom_pages_visibility(self):
        if self._page_range_var.get() == PAGE_RANGE_CUSTOM:
            if not self._custom_pages_visible:
                try:
                    self._pg_custom_row.grid()
                except Exception:
                    pass
                if not getattr(self, "_custom_pages_packed", False):
                    self.custom_pages_entry.pack(fill="x")
                    self._custom_pages_packed = True
                self._custom_pages_visible = True
        else:
            if self._custom_pages_visible:
                try:
                    self._pg_custom_row.grid_remove()
                except Exception:
                    pass
                self._custom_pages_visible = False

    def _on_page_range_change(self, _v=None):
        self._refresh_custom_pages_visibility()
        if getattr(self, "_syncing_page_ui", False):
            return
        mode = PAGE_MODE_BY_LABEL.get(self._page_range_var.get(), "all")
        indices = self._selected_indices()
        if not indices:
            self._save_config()
            return
        editable, skipped = self._editable_indices(indices)
        if mode == "custom":
            # Chờ nhập ô tùy chọn rồi Enter — không áp vội giá trị cũ
            self._save_config()
            return
        for i in editable:
            self.file_list[i].page_mode = mode
            self.file_list[i].page_range_text = ""
        if editable:
            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._log(f"Đã đặt trang in [{self._page_range_var.get()}] cho {len(editable)} tệp tin")
        if skipped:
            self._log(f"⚠️ Bỏ qua {len(skipped)} tệp đang in / đã in xong khi đặt trang in")
        self._save_config()

    def _commit_custom_pages(self):
        """Validate + apply custom range text to selected editable files."""
        if getattr(self, "_syncing_page_ui", False):
            return
        if self._page_range_var.get() != PAGE_RANGE_CUSTOM:
            return
        raw = self._custom_pages_var.get().strip()
        indices = self._selected_indices()
        if not indices:
            return
        editable, skipped = self._editable_indices(indices)
        if not editable:
            if skipped:
                messagebox.showwarning("Không thể chỉnh sửa", "Các tệp đang in / đã in xong, không thể đổi trang in.")
            return
        errors = []
        for i in editable:
            f = self.file_list[i]
            try:
                resolve_page_selection("custom", raw, f.page_count)
            except ValueError as exc:
                errors.append(f"{f.filename}: {exc}")
        if errors:
            messagebox.showwarning("Trang in không hợp lệ", "\n".join(errors[:5]))
            return
        for i in editable:
            self.file_list[i].page_mode = "custom"
            self.file_list[i].page_range_text = raw
        self._refresh_tree(preserve_selection=True)
        self._update_stats_pill()
        self._log(f"Đã đặt trang in tùy chọn [{raw}] cho {len(editable)} tệp tin")
        if skipped:
            self._log(f"⚠️ Bỏ qua {len(skipped)} tệp đang in / đã in xong")
        self._save_config()

    def _apply_pages_to_all(self):
        """Apply current panel page setting to every editable file (spec §9)."""
        mode = PAGE_MODE_BY_LABEL.get(self._page_range_var.get(), "all")
        raw = self._custom_pages_var.get().strip()
        if mode == "custom":
            if not raw:
                messagebox.showinfo("Áp dụng cho tất cả", "Vui lòng nhập phạm vi trang tùy chọn trước.")
                return
            # Validate mẫu trên file đầu để báo lỗi sớm (số trang mỗi file check lúc in)
            try:
                parse_page_range(raw, 10 ** 9)
            except ValueError as exc:
                messagebox.showwarning("Trang in không hợp lệ", str(exc))
                return
        updated, skipped = 0, 0
        for f in self.file_list:
            if is_edit_locked(f.status):
                skipped += 1
                continue
            f.page_mode = mode
            f.page_range_text = raw if mode == "custom" else ""
            updated += 1
        self._refresh_tree(preserve_selection=True)
        self._update_stats_pill()
        msg = f"Đã áp dụng trang in [{self._page_range_var.get()}{' ' + raw if mode == 'custom' else ''}] cho {updated} tệp tin"
        if skipped:
            msg += f" (bỏ qua {skipped} tệp đang in/đã in)"
        self._log(msg)
        self._save_config()

    # ── Custom paper size ────────────────────────────────────────────

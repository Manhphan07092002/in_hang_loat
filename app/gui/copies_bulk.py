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


class CopiesBulkMixin:

    def _on_tree_right_click(self, event):
        """Right-click popup context menu for fast copy modifications."""
        item = self.tree.identify_row(event.y)
        if item and item not in self.tree.selection():
            self.tree.selection_set(item)

        sel = self.tree.selection()
        if not sel:
            return

        menu = tk.Menu(self, tearoff=0, font=("Segoe UI", 10))
        n = len(sel)
        suffix = f" ({n} tệp)" if n > 1 else ""

        menu.add_command(label="🔍 Xem trước chi tiết (Mở cửa sổ lớn)...", command=self._preview_selected)
        menu.add_separator()
        menu.add_command(label=f"✏️ Sửa số bản in...{suffix}", command=self._edit_selected_copies)
        if n == 1:
            menu.add_command(label="📄 Đặt trang in riêng...",
                             command=lambda: self._start_pages_inline_edit(sel[0]))
        menu.add_separator()
        menu.add_command(label="🎯 In các mục đang chọn", command=self.start_print_selected)
        menu.add_separator()
        menu.add_command(label="➕ Tăng +1 bản (+)", command=lambda: self._change_selected_copies(1))
        menu.add_command(label="➖ Giảm -1 bản (-)", command=lambda: self._change_selected_copies(-1))

        # Preset submenu
        preset_menu = tk.Menu(menu, tearoff=0, font=("Segoe UI", 10))
        for count in [1, 2, 3, 4, 5, 10, 20]:
            preset_menu.add_command(
                label=f"{count} bản",
                command=lambda c=count: self._set_selected_copies(c),
            )
        menu.add_cascade(label="📋 Đặt nhanh số bản", menu=preset_menu)

        menu.add_separator()
        menu.add_command(label="▲ Di chuyển lên (Ctrl+Up)", command=self.move_selected_up)
        menu.add_command(label="▼ Di chuyển xuống (Ctrl+Down)", command=self.move_selected_down)
        menu.add_separator()
        menu.add_command(label="💾 Lưu danh sách (JSON)...", command=self.save_queue_session)
        menu.add_command(label="📂 Mở danh sách (JSON)...", command=self.load_queue_session)
        menu.add_separator()
        menu.add_command(label="🗑️ Xóa khỏi hàng đợi (Del)", command=self.remove_selected)

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _change_selected_copies(self, delta: int):
        sel = self.tree.selection()
        if not sel:
            return
        last_val = 1
        skipped_names = []
        updated_count = 0
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    f = self.file_list[idx]
                    if f.status in (FileStatus.PRINTING, FileStatus.PRINTED):
                        skipped_names.append(f.filename)
                        continue
                    new_val = max(1, min(999, f.copies + delta))
                    f.copies = new_val
                    last_val = new_val
                    updated_count += 1
            except Exception:
                pass

        if skipped_names:
            self._log(f"⚠️ Bỏ qua {len(skipped_names)} tệp đang in / đã in xong khi chỉnh số bản")
            if updated_count == 0:
                messagebox.showwarning(
                    "Không thể chỉnh sửa",
                    "Các tệp tin được chọn đang trong quá trình in hoặc đã hoàn thành in,\nkhông thể sửa đổi số bản in.",
                )

        if updated_count > 0:
            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._copies_var.set(last_val)

    def _set_selected_copies(self, copies: int):
        sel = self.tree.selection()
        if not sel:
            return
        copies = max(1, min(999, copies))
        skipped_names = []
        updated_count = 0
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    f = self.file_list[idx]
                    if f.status in (FileStatus.PRINTING, FileStatus.PRINTED):
                        skipped_names.append(f.filename)
                        continue
                    f.copies = copies
                    updated_count += 1
            except Exception:
                pass

        if skipped_names:
            self._log(f"⚠️ Bỏ qua {len(skipped_names)} tệp đang in / đã in xong khi đặt số bản")
            if updated_count == 0:
                messagebox.showwarning(
                    "Không thể chỉnh sửa",
                    "Các tệp tin được chọn đang trong quá trình in hoặc đã hoàn thành in,\nkhông thể sửa đổi số bản in.",
                )

        if updated_count > 0:
            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._copies_var.set(copies)
            self._log(f"Đã đặt {copies} bản in cho {updated_count} tệp tin")

    def _prompt_change_copies_dialog(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Thông báo", "Vui lòng chọn ít nhất một tệp tin trong danh sách để đổi số bản in.")
            return

        valid_items = []
        skipped_names = []
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    f = self.file_list[idx]
                    if f.status in (FileStatus.PRINTING, FileStatus.PRINTED):
                        skipped_names.append(f.filename)
                    else:
                        valid_items.append(f)
            except Exception:
                pass

        if not valid_items:
            messagebox.showwarning(
                "Không thể chỉnh sửa",
                "Tất cả các tệp tin được chọn đều đang trong quá trình in hoặc đã in xong,\nkhông thể sửa đổi số bản in.",
            )
            return

        current_val = valid_items[0].copies
        count_text = f"{len(valid_items)} tệp tin có thể chỉnh sửa" if len(valid_items) > 1 else f"tệp '{valid_items[0].filename}'"
        dialog = ctk.CTkInputDialog(
            text=f"Nhập số bản in mới cho {count_text} (1 - 999):",
            title="Chỉnh Sửa Số Bản In",
        )
        val_str = dialog.get_input()
        if val_str is not None and val_str.strip():
            try:
                val = int(val_str.strip())
                if 1 <= val <= 999:
                    self._set_selected_copies(val)
                else:
                    messagebox.showwarning("Cảnh báo", "Số bản in phải từ 1 đến 999.")
            except ValueError:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập số nguyên hợp lệ.")

    def _inc_copies(self):
        try:
            v = int(self._copies_var.get())
        except Exception:
            v = 1
        if v < 999:
            new_v = v + 1
            self._copies_var.set(new_v)
            if self.tree.selection():
                self._set_selected_copies(new_v)
            self._save_config()

    def _dec_copies(self):
        try:
            v = int(self._copies_var.get())
        except Exception:
            v = 1
        if v > 1:
            new_v = v - 1
            self._copies_var.set(new_v)
            if self.tree.selection():
                self._set_selected_copies(new_v)
            self._save_config()

    def _on_copies_entry_enter(self, _event=None):
        try:
            val = int(self._copies_var.get())
            val = max(1, min(999, val))
            self._copies_var.set(val)
            if self.tree.selection():
                self._set_selected_copies(val)
            self._save_config()
        except Exception:
            pass

    def _apply_copies_to_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Thông báo", "Vui lòng chọn các tệp tin trong hàng đợi trước khi áp dụng.")
            return
        try:
            val = int(self._copies_var.get())
            val = max(1, min(999, val))
            self._copies_var.set(val)
            self._set_selected_copies(val)
            self._save_config()
        except Exception:
            pass

    def _apply_default_copies_to_all(self):
        try:
            c = max(1, min(999, int(self._default_copies_var.get())))
        except Exception:
            c = 1
        self._default_copies_var.set(c)
        updated = 0
        skipped = 0
        for f in self.file_list:
            if f.status in (FileStatus.PRINTING, FileStatus.PRINTED):
                skipped += 1
                continue
            f.copies = c
            updated += 1
        self._refresh_tree(preserve_selection=True)
        self._update_stats_pill()
        self._update_queue_summary()
        msg = f"Đã áp dụng số bản mặc định ({c} bản) cho {updated} tệp tin"
        if skipped > 0:
            msg += f" (Bỏ qua {skipped} tệp đang in/đã in)"
        self._log(msg)
        self._save_config()

    def _selected_indices(self) -> list[int]:
        """Queue indices from current tree selection (in selection order)."""
        indices = []
        for item in self.tree.selection():
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    indices.append(idx)
            except Exception:
                pass
        return indices

    def _editable_indices(self, indices: list[int]) -> tuple[list[int], list[str]]:
        """Split into editable indices + skipped filenames (printing/printed)."""
        ok, skipped = [], []
        for i in indices:
            f = self.file_list[i]
            if is_edit_locked(f.status):
                skipped.append(f.filename)
            else:
                ok.append(i)
        return ok, skipped

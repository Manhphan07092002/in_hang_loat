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


class CopiesEditMixin:

    def _on_tree_click(self, event):
        """Single click on copies column (#7) edits copies; on Trang In (#6) edits pages."""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            if col == "#7":
                item = self.tree.identify_row(event.y)
                if item:
                    self.after(40, lambda: self._start_inline_edit(item, "#7"))
            elif col == "#6":
                item = self.tree.identify_row(event.y)
                if item:
                    self.after(40, lambda: self._start_pages_inline_edit(item))

    def _on_tree_double_click(self, event):
        """Double click on copies (#7) edits copies, on Trang In (#6) edits pages, elsewhere opens preview."""
        region = self.tree.identify("region", event.x, event.y)
        if region in ("cell", "tree"):
            item = self.tree.identify_row(event.y)
            if item:
                col = self.tree.identify_column(event.x)
                if col == "#7":
                    self._start_inline_edit(item, "#7")
                elif col == "#6":
                    self._start_pages_inline_edit(item)
                else:
                    self._preview_selected()

    def _start_inline_edit(self, item, col_id="#7"):
        """Spawn an in-place entry widget on the treeview cell for editing copies."""
        if not item:
            return
        bbox = self.tree.bbox(item, col_id)
        if not bbox:
            return

        x, y, w, h = bbox
        try:
            values = self.tree.item(item, "values")
            idx = int(values[0]) - 1
        except Exception:
            return

        if idx < 0 or idx >= len(self.file_list):
            return

        f = self.file_list[idx]
        # Check if file is currently printing or already printed
        if f.status == FileStatus.PRINTING:
            messagebox.showwarning(
                "Không thể chỉnh sửa",
                f"Tệp tin '{f.filename}' đang trong quá trình in ấn.\nBạn không thể sửa số bản in lúc này.",
            )
            return
        if f.status == FileStatus.PRINTED:
            messagebox.showwarning(
                "Không thể chỉnh sửa",
                f"Tệp tin '{f.filename}' đã hoàn thành in.\nKhông thể sửa đổi số bản in của tệp đã in.",
            )
            return

        if self._edit_entry is not None:
            try:
                self._edit_entry.destroy()
            except Exception:
                pass
            self._edit_entry = None

        entry = tk.Entry(
            self.tree,
            justify="center",
            font=("Segoe UI", 10, "bold"),
            bd=2,
            relief="solid",
            highlightthickness=1,
            highlightcolor="#2563EB",
            bg="#EFF6FF",
            fg="#1E40AF",
        )
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, str(f.copies))
        entry.select_range(0, "end")
        entry.focus_set()
        self._edit_entry = entry

        def _commit_and_move(direction=0):
            try:
                val = int(entry.get().strip())
                val = max(1, min(999, val))
            except (ValueError, AttributeError):
                val = self.file_list[idx].copies

            self.file_list[idx].copies = val
            if self._edit_entry == entry:
                entry.destroy()
                self._edit_entry = None

            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._copies_var.set(val)

            # Fast navigation to next/previous available (unprinted) row
            if direction != 0:
                children = self.tree.get_children()
                try:
                    curr_pos = children.index(item)
                    next_pos = curr_pos + direction
                    while 0 <= next_pos < len(children):
                        cand_item = children[next_pos]
                        cand_idx = int(self.tree.item(cand_item, "values")[0]) - 1
                        if 0 <= cand_idx < len(self.file_list):
                            cand_f = self.file_list[cand_idx]
                            if cand_f.status not in (FileStatus.PRINTING, FileStatus.PRINTED):
                                self.tree.selection_set(cand_item)
                                self.tree.see(cand_item)
                                self.after(30, lambda c=cand_item: self._start_inline_edit(c, col_id))
                                break
                        next_pos += direction
                except Exception:
                    pass

        def _cancel(_e=None):
            if self._edit_entry == entry:
                entry.destroy()
                self._edit_entry = None

        entry.bind("<Return>", lambda e: _commit_and_move(1))
        entry.bind("<Tab>", lambda e: (_commit_and_move(1), "break")[1])
        entry.bind("<Shift-Tab>", lambda e: (_commit_and_move(-1), "break")[1])
        entry.bind("<Down>", lambda e: _commit_and_move(1))
        entry.bind("<Up>", lambda e: _commit_and_move(-1))
        entry.bind("<Escape>", _cancel)
        entry.bind("<FocusOut>", lambda e: _commit_and_move(0))

    def _edit_selected_copies(self):
        """Dedicated edit button handler — verifies printing state and triggers editor."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Chỉnh sửa số bản in", "Vui lòng chọn ít nhất một tệp tin trong danh sách để sửa số bản in.")
            return

        if len(sel) == 1:
            item = sel[0]
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
            except Exception:
                return
            if 0 <= idx < len(self.file_list):
                f = self.file_list[idx]
                if f.status == FileStatus.PRINTING:
                    messagebox.showwarning(
                        "Không thể chỉnh sửa",
                        f"Tệp tin '{f.filename}' đang trong quá trình in ấn.\nBạn không thể sửa số bản in lúc này.",
                    )
                    return
                if f.status == FileStatus.PRINTED:
                    messagebox.showwarning(
                        "Không thể chỉnh sửa",
                        f"Tệp tin '{f.filename}' đã in xong.\nKhông thể sửa đổi số bản in của tệp đã in.",
                    )
                    return
                self._start_inline_edit(item, "#7")
        else:
            self._prompt_change_copies_dialog()

    def _on_tree_keypress(self, event):
        """Handle quick keyboard shortcuts for adjusting copies on selected items."""
        if self._edit_entry is not None:
            return

        sel = self.tree.selection()
        if not sel:
            return

        # Plus / Equal key -> Increase copies by 1
        if event.char in ("+", "="):
            self._change_selected_copies(delta=1)
            return "break"
        # Minus / Underscore key -> Decrease copies by 1
        elif event.char in ("-", "_"):
            self._change_selected_copies(delta=-1)
            return "break"
        # Number keys 1-9 -> Set exact copy count
        elif event.char in "123456789":
            copies = int(event.char)
            self._set_selected_copies(copies)
            return "break"
        # F2, Return, Space -> Open inline editor for selected item
        elif event.keysym in ("F2", "Return", "space"):
            self._edit_selected_copies()
            return "break"

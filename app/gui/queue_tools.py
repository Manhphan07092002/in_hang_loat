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


class QueueToolsMixin:

    def _on_toolbar_resize(self, _event=None):
        """Debounce toolbar reflow on width change (skip if width unchanged)."""
        try:
            w = self._toolbar.winfo_width()
            if w == getattr(self, "_toolbar_last_w", None):
                return
            self._toolbar_last_w = w
            if self._reflow_job is not None:
                self.after_cancel(self._reflow_job)
        except Exception:
            pass
        try:
            self._reflow_job = self.after(120, self._reflow_toolbar)
        except Exception:
            pass

    def _reflow_toolbar(self):
        """Xếp 9 nút + checkbox vào các dòng grid vừa khung (tham lam theo
        chiều rộng thực). Không nút nào bị cắt/chồng; thứ tự đọc giữ nguyên.
        """
        self._reflow_job = None
        if getattr(self, "_reflowing", False):
            return
        if not getattr(self, "_is_alive", False):
            return
        self._reflowing = True
        try:
            flow = self._toolbar_flow
            try:
                avail = max(220, flow.winfo_width() or 0)
            except Exception:
                return
            if avail <= 0:
                return

            def _w(w):
                try:
                    return w.winfo_reqwidth() + 5
                except Exception:
                    return 80

            rows: list[list] = []
            cur: list = []
            curw = 0
            for b in self._toolbar_buttons:
                bw = _w(b)
                if cur and curw + bw > avail:
                    rows.append(cur)
                    cur, curw = [], 0
                cur.append(b)
                curw += bw
            if cur:
                rows.append(cur)

            for b in self._toolbar_buttons:
                try:
                    b.grid_forget()
                except Exception:
                    pass
            try:
                self._subfolder_check.grid_forget()
            except Exception:
                pass
            for r, rowbtns in enumerate(rows):
                for c, b in enumerate(rowbtns):
                    b.grid(row=r, column=c, padx=2, pady=2, sticky="w")
            # Checkbox luôn cuối dòng cuối, dính phải
            flow.grid_columnconfigure(len(rows[-1]) if rows else 0, weight=1)
            self._subfolder_check.grid(row=len(rows) - 1, column=len(rows[-1]),
                                       padx=(8, 0), pady=2, sticky="e")
            # Cập nhật membership để test/inspect biết nút đang dòng nào
            self._toolbar_row1_btns = list(rows[0]) if rows else []
            self._toolbar_row2_btns = [b for r in rows[1:] for b in r]
        finally:
            self._reflowing = False

    def _on_tree_hover(self, event):
        """Show full filename tooltip when hovering a truncated filename cell."""
        try:
            if self._tree_tip_job is not None:
                self.after_cancel(self._tree_tip_job)
                self._tree_tip_job = None
        except Exception:
            pass
        self._hide_tree_tooltip()

        def _show():
            try:
                if self.tree.identify("region", event.x, event.y) != "cell":
                    return
                if self.tree.identify_column(event.x) != "#2":  # filename column
                    return
                item = self.tree.identify_row(event.y)
                if not item:
                    return
                idx = int(self.tree.item(item, "values")[0]) - 1
                if not (0 <= idx < len(self.file_list)):
                    return
                full = self.file_list[idx].filename
                # Chỉ hiện khi text thực sự bị cắt (đo rộng hơn cột)
                from tkinter import font as tkfont
                w = tkfont.Font(family="Segoe UI", size=10).measure(full)
                if w <= self.tree.column("filename", option="width") - 6:
                    return
                tip = tk.Toplevel(self)
                tip.wm_overrideredirect(True)
                tip.wm_geometry(f"+{event.x_root + 14}+{event.y_root + 12}")
                lbl = tk.Label(tip, text=full, bg="#0F172A", fg="#F8FAFC",
                               font=("Segoe UI", 10), padx=8, pady=4,
                               wraplength=420, justify="left")
                lbl.pack()
                self._tree_tip = tip
            except Exception:
                pass

        try:
            self._tree_tip_job = self.after(450, _show)
        except Exception:
            pass

    def _hide_tree_tooltip(self):
        try:
            if self._tree_tip is not None:
                self._tree_tip.destroy()
        except Exception:
            pass
        self._tree_tip = None

    def _select_all_rows(self, _e=None):
        self.tree.selection_set(self.tree.get_children())
        return "break"

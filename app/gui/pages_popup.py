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


class PagesPopupMixin:

    def _start_pages_inline_edit(self, item):
        """Popup editor for the 'Trang In' cell: mode dropdown + custom range entry."""
        if not item:
            return
        try:
            values = self.tree.item(item, "values")
            idx = int(values[0]) - 1
        except Exception:
            return
        if idx < 0 or idx >= len(self.file_list):
            return
        f = self.file_list[idx]
        if is_edit_locked(f.status):
            messagebox.showwarning(
                "Không thể chỉnh sửa",
                f"Tệp tin '{f.filename}' đang trong quá trình in hoặc đã in xong.\nKhông thể đổi trang in lúc này.",
            )
            return

        # Close any stale copies editor
        if self._edit_entry is not None:
            try:
                self._edit_entry.destroy()
            except Exception:
                pass
            self._edit_entry = None

        # Position popup near the clicked cell
        try:
            bbox = self.tree.bbox(item, "#6")
            if bbox:
                x, y, _w, h = bbox
                px = self.tree.winfo_rootx() + x
                py = self.tree.winfo_rooty() + y + h
            else:
                px, py = self.tree.winfo_rootx() + 100, self.tree.winfo_rooty() + 60
        except Exception:
            px, py = self.tree.winfo_rootx() + 100, self.tree.winfo_rooty() + 60

        dlg = ctk.CTkToplevel(self)
        dlg.title(f"Trang in — {f.filename}")
        dlg.geometry(f"320x210+{px}+{py}")
        dlg.transient(self)
        dlg.resizable(False, False)
        try:
            dlg.grab_set()
        except Exception:
            pass

        ctk.CTkLabel(
            dlg, text=f"📄 {f.filename}"[:42],
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(padx=14, pady=(12, 2), anchor="w")
        ctk.CTkLabel(
            dlg, text=f"Tài liệu có {f.page_count} trang. Chọn chế độ in:",
            font=ctk.CTkFont(size=11),
            text_color=THEME_COLORS["text_muted"],
        ).pack(padx=14, pady=(0, 8), anchor="w")

        mode_var = tk.StringVar(value=PAGE_LABEL_BY_MODE.get(f.page_mode, PAGE_RANGE_ALL))
        val_var = tk.StringVar(value=f.page_range_text)

        combo = ctk.CTkComboBox(
            dlg, values=PAGE_RANGE_OPTIONS, variable=mode_var,
            width=292, height=30, state="readonly",
            font=ctk.CTkFont(size=11),
        )
        combo.pack(padx=14, pady=(0, 6))

        entry = ctk.CTkEntry(
            dlg, textvariable=val_var, width=292, height=30,
            placeholder_text="Ví dụ: 1,3,5-8,12",
            font=ctk.CTkFont(size=11),
        )

        def _toggle_entry(_v=None):
            if mode_var.get() == PAGE_RANGE_CUSTOM:
                entry.pack(padx=14, pady=(0, 6))
                entry.focus_set()
            else:
                try:
                    entry.pack_forget()
                except Exception:
                    pass

        combo.configure(command=_toggle_entry)
        _toggle_entry()

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(padx=14, pady=(4, 12), fill="x")

        def _on_ok():
            mode = PAGE_MODE_BY_LABEL.get(mode_var.get(), "all")
            raw = val_var.get().strip()
            if mode == "custom":
                try:
                    resolve_page_selection("custom", raw, f.page_count)
                except ValueError as exc:
                    messagebox.showwarning("Trang in không hợp lệ",
                                           f"{f.filename} ({f.page_count} trang):\n{exc}",
                                           parent=dlg)
                    return
            f.page_mode = mode
            f.page_range_text = raw if mode == "custom" else ""
            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._sync_page_ui_from_selection()
            self._log(f"Đã đặt trang in [{f.pages_display()}] cho '{f.filename}'")
            self._save_config()
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btn_row, text="✔ Đồng ý", command=_on_ok, height=30,
            fg_color=THEME_COLORS["primary"],
            hover_color=THEME_COLORS["primary_hover"],
            text_color="#FFFFFF", text_color_disabled="#FFFFFF",
            font=ctk.CTkFont(size=11, weight="bold"),
        ).pack(side="left", expand=True, fill="x", padx=(0, 4))
        ctk.CTkButton(
            btn_row, text="Hủy", command=dlg.destroy, height=30,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            text_color_disabled=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(size=11),
        ).pack(side="left", expand=True, fill="x", padx=(4, 0))

        dlg.bind("<Escape>", lambda e: dlg.destroy())
        entry.bind("<Return>", lambda e: _on_ok())

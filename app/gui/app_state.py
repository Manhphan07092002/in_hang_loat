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


class AppStateMixin:

    def _load_config(self) -> dict:
        return load_config()

    def _save_config(self):
        try:
            printer_name = self._printer_var.get()
            if printer_name.startswith("("):
                printer_name = ""
            cfg = {
                "printer": printer_name,
                "duplex": self._duplex_var.get(),
                "paper": self._resolve_gui_paper(),
                "custom_paper_w": self._custom_paper_w.get() if hasattr(self, "_custom_paper_w") else "210",
                "custom_paper_h": self._custom_paper_h.get() if hasattr(self, "_custom_paper_h") else "297",
                "orientation": self._orient_var.get(),
                "copies": self._copies_var.get(),
                "fit_to_page": self._fit_to_page.get(),
                "include_subfolders": self._include_subfolders.get(),
                "theme": "Dark" if self.theme_switch.get() == 1 else "Light",
                "remove_blanks": self._remove_blanks_var.get(),
                "separator_sheet": self._separator_sheet_var.get(),
                "reverse_order": self._reverse_order_var.get(),
                "binding_margin": self._binding_margin_var.get(),
                "failover_printer": self._failover_printer_var.get(),
                "invoice_mode": self._invoice_mode_var.get() if hasattr(self, "_invoice_mode_var") else INVOICE_AUTO,
                "update_last_check": self._cfg.get("update_last_check", ""),
                "update_skip_version": self._cfg.get("update_skip_version", ""),
            }
            save_config(cfg)
        except Exception as exc:
            print(f"Error saving config: {exc}")

    # ── Drag & Drop Handling ─────────────────────────────────────────

    def _init_drag_and_drop(self):
        try:
            import windnd
            windnd.hook_dropfiles(self, func=self._on_files_dropped)
        except Exception as exc:
            print(f"Could not initialize windnd drag-and-drop: {exc}")

    def _on_files_dropped(self, files):
        if not getattr(self, "_is_alive", False):
            return
        decoded_paths = []
        for item in files:
            if isinstance(item, bytes):
                try:
                    p = item.decode("utf-8")
                except UnicodeDecodeError:
                    try:
                        p = item.decode("mbcs")
                    except Exception:
                        p = str(item)
            else:
                p = str(item)
            decoded_paths.append(p)

        all_paths_to_add = []
        recursive = self._include_subfolders.get()

        for p in decoded_paths:
            if not os.path.exists(p):
                continue
            if os.path.isdir(p):
                if recursive:
                    for root, _dirs, fnames in os.walk(p):
                        for fn in fnames:
                            ext = os.path.splitext(fn)[1].lower()
                            if ext in SUPPORTED_EXTENSIONS:
                                all_paths_to_add.append(os.path.join(root, fn))
                else:
                    for fn in os.listdir(p):
                        fp = os.path.join(p, fn)
                        if os.path.isfile(fp):
                            ext = os.path.splitext(fn)[1].lower()
                            if ext in SUPPORTED_EXTENSIONS:
                                all_paths_to_add.append(fp)
            elif os.path.isfile(p):
                ext = os.path.splitext(p)[1].lower()
                if ext in SUPPORTED_EXTENSIONS:
                    all_paths_to_add.append(p)

        if all_paths_to_add:
            self._async_queue.put((self._add_to_queue, (all_paths_to_add,)))
        else:
            self._async_queue.put((self._log, ("⚠️ Không tìm thấy tệp tin được hỗ trợ trong các mục vừa kéo thả",)))

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


class SettingsPaperMixin:

    def _on_paper_change(self, _v=None):
        is_custom = self._paper_var.get() == CUSTOM_PAPER_LABEL
        try:
            if is_custom:
                self.custom_paper_frame.pack(fill="x", pady=(6, 0))
            else:
                self.custom_paper_frame.pack_forget()
        except Exception:
            pass
        self._save_config()

    def _on_custom_paper_typed(self):
        try:
            w = self._custom_paper_w.get().strip()
            h = self._custom_paper_h.get().strip()
            parse_custom_paper(f"{w}x{h}")
            self._save_config()
        except Exception:
            pass

    def _resolve_gui_paper(self) -> str:
        """Return standard key or 'WxH' string for the print pipeline."""
        if self._paper_var.get() == CUSTOM_PAPER_LABEL:
            try:
                w = self._custom_paper_w.get().strip()
                h = self._custom_paper_h.get().strip()
                parse_custom_paper(f"{w}x{h}")
                return f"{w}x{h}"
            except Exception:
                return "A4"
        return self._paper_var.get()

    def _sync_paper_from_config(self):
        """Restore custom W×H into entries when config holds 'WxH'."""
        saved = self._cfg.get("paper", "A4")
        if saved in PAPER_SIZES or saved == CUSTOM_PAPER_LABEL:
            self._paper_var.set(saved)
        else:
            try:
                w, h = parse_custom_paper(saved)
                self._custom_paper_w.set(str(int(w) if float(w).is_integer() else w))
                self._custom_paper_h.set(str(int(h) if float(h).is_integer() else h))
                self._paper_var.set(CUSTOM_PAPER_LABEL)
            except Exception:
                self._paper_var.set("A4")
        self._on_paper_change()

    # ═══════════════════════════════════════════════════════════════════
    #  PRINTING WORKFLOW
    # ═══════════════════════════════════════════════════════════════════

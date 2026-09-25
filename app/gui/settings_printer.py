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


class SettingsPrinterMixin:

    def _build_sec_printer(self, card, _section, _row):
        """Mục Máy in (tách từ _build_settings_card)."""
        # ══ Section 1: Máy in ══════════════════════════════════════
        _section("🖨️  MÁY IN")
        p_box = _row()
        p_box.grid_columnconfigure(0, weight=1)
        p_box.grid_columnconfigure(1, weight=0)

        self.printer_combo = ctk.CTkComboBox(
            p_box, variable=self._printer_var,
            state="readonly", command=self._on_printer_change,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=8, height=30,
        )
        self.printer_combo.grid(row=0, column=0, sticky="ew", padx=(0, 6))

        ctk.CTkButton(
            p_box, text="⟳ Làm mới", width=80, height=30,
            command=self.refresh_printers,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=8,
        ).grid(row=0, column=1)

        self.printer_status_lbl = ctk.CTkLabel(
            card, text="", font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=THEME_COLORS["text_muted"],
        )
        self.printer_status_lbl.grid(row=_section.row, column=0, sticky="w", padx=14, pady=(0, 2))
        _section.row += 1

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


class SettingsShellMixin:

    def _build_settings_card(self, parent):
        card = ctk.CTkScrollableFrame(
            parent, fg_color=THEME_COLORS["card"],
            corner_radius=12, border_width=1,
            border_color=THEME_COLORS["card_border"],
        )
        card.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
        card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            card, text="⚙️ Cài Đặt In Ấn & Tính Năng Mở Rộng",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=THEME_COLORS["text"],
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(10, 4))

        def _section(title: str):
            """Full-width section header with divider. Returns next free row."""
            r = _section.row
            ctk.CTkLabel(
                card, text=title,
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                text_color=THEME_COLORS["primary"][0] if ctk.get_appearance_mode() == "Light" else THEME_COLORS["primary"][1],
            ).grid(row=r, column=0, sticky="w", padx=14, pady=(8, 2))
            _section.row = r + 1
            return r + 1
        _section.row = 1

        def _row():
            f = ctk.CTkFrame(card, fg_color="transparent")
            f.grid(row=_section.row, column=0, sticky="ew", padx=14, pady=2)
            f.grid_columnconfigure(0, weight=1)
            _section.row += 1
            return f

        self._build_sec_printer(card, _section, _row)
        self._build_sec_copies(card, _section, _row)

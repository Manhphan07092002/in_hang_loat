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


class SettingsExtraMixin:

    def _build_sec_invoice(self, card, _section, _row):
        """Mục Hóa đơn điện tử (tách từ _build_settings_card)."""
        # ══ Section: Hóa đơn điện tử ══════════════════════════════
        _section("🧾  HÓA ĐƠN ĐIỆN TỬ")
        inv_row = _row()
        for text in INVOICE_MODES:
            ctk.CTkRadioButton(
                inv_row, text=text, variable=self._invoice_mode_var, value=text,
                command=self._on_invoice_mode_change,
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color=THEME_COLORS["text"],
                radiobutton_width=16, radiobutton_height=16,
            ).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(
            card, text="Tự động: hóa đơn ≥2 trang → 2 mặt cạnh dài • Cảnh báo: hỏi trước khi áp",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=THEME_COLORS["text_muted"], wraplength=300, justify="left",
        ).grid(row=_section.row, column=0, sticky="w", padx=14, pady=(0, 2))
        _section.row += 1

    def _build_sec_extra(self, card, _section, _row):
        """Mục Tùy chọn + In song song (tách từ _build_settings_card)."""
        # ══ Section: Tùy chọn thông minh ═════════════════════════
        _section("✨  TÙY CHỌN")
        smart_frame = ctk.CTkFrame(card, fg_color=THEME_COLORS["card_alt"], corner_radius=8)
        smart_frame.grid(row=_section.row, column=0, sticky="ew", padx=14, pady=2)
        _section.row += 1
        smart_frame.grid_columnconfigure(0, weight=1)

        for _r, (_txt, _var) in enumerate([
            ("🚫 Tự động bỏ trang trắng", self._remove_blanks_var),
            ("🔄 In đảo ngược (Cuối → 1)", self._reverse_order_var),
            ("📄 Chèn tờ bìa phân cách", self._separator_sheet_var),
            ("📐 Vừa trang giấy (Fit to page)", self._fit_to_page),
        ]):
            ctk.CTkCheckBox(
                smart_frame, text=_txt,
                variable=_var,
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color=THEME_COLORS["text"],
                checkbox_width=16, checkbox_height=16, corner_radius=4,
            ).grid(row=_r, column=0, sticky="w", padx=8, pady=2)

        # ── In song song ─────────────────────────────────────────────
        parallel_bar = ctk.CTkFrame(card, fg_color="transparent")
        parallel_bar.grid(row=_section.row, column=0, sticky="ew", padx=14, pady=(4, 10))
        _section.row += 1

        ctk.CTkButton(
            parallel_bar, text="🖨️ In song song nhiều máy in...",
            command=self._open_parallel_printers_dialog,
            height=28, fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            corner_radius=6,
        ).pack(side="left")

        self.parallel_status_lbl = ctk.CTkLabel(
            parallel_bar, text="",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=THEME_COLORS["success"][0],
        )
        self.parallel_status_lbl.pack(side="left", padx=(8, 0))

        self.duplex_warn_lbl = ctk.CTkLabel(
            parallel_bar, text="",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=THEME_COLORS["warning"][0],
        )
        self.duplex_warn_lbl.pack(side="right")

        self._sync_paper_from_config()

    # ═══════════════════════════════════════════════════════════════════
    #  RIGHT COLUMN: PROGRESS & CONTROL ACTIONS
    # ═══════════════════════════════════════════════════════════════════

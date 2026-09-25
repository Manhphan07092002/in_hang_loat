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


class SettingsPagesMixin:

    def _build_sec_copies(self, card, _section, _row):
        """Mục Tệp đang chọn: tên file, stepper, mặc định, trang in, nhận diện."""
        # ══ Section 2: Tệp đang chọn (số bản + trang in) ═══════════
        _section("📄  TỆP ĐANG CHỌN")

        # File name — full width, truncated
        self.selected_file_lbl = ctk.CTkLabel(
            card, text="📄 (Chưa chọn tệp)",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=THEME_COLORS["text"],
            anchor="w",
        )
        self.selected_file_lbl.grid(row=_section.row, column=0, sticky="ew", padx=14, pady=(0, 2))
        _section.row += 1

        # Copies stepper
        c_box = _row()
        ctk.CTkLabel(
            c_box, text="Số bản in:",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=THEME_COLORS["text"],
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            c_box, text="−", width=30, height=28,
            command=self._dec_copies,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=6,
        ).pack(side="left")

        self.copies_entry = ctk.CTkEntry(
            c_box, textvariable=self._copies_var,
            width=52, height=28, justify="center",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.copies_entry.pack(side="left", padx=4)
        self.copies_entry.bind("<Return>", lambda e: self._on_copies_entry_enter())
        self.copies_entry.bind("<FocusOut>", lambda e: self._on_copies_entry_enter())

        ctk.CTkButton(
            c_box, text="+", width=30, height=28,
            command=self._inc_copies,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(size=14, weight="bold"),
            corner_radius=6,
        ).pack(side="left")

        ctk.CTkButton(
            c_box, text="✏️ Sửa", width=56, height=28,
            command=self._edit_selected_copies,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        ).pack(side="left", padx=(4, 0))

        # Page mode + custom entry (full-width rows)
        pg_mode_row = _row()
        self.page_range_combo = ctk.CTkComboBox(
            pg_mode_row, values=PAGE_RANGE_OPTIONS, variable=self._page_range_var,
            height=30, state="readonly",
            command=self._on_page_range_change,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=8,
        )
        self.page_range_combo.pack(fill="x")

        pg_custom_row = _row()
        self._pg_custom_row = pg_custom_row
        pg_custom_row.grid_remove()  # CTkFrame trống mặc định cao 200px → ẩn hẳn tới khi cần
        self.custom_pages_entry = ctk.CTkEntry(
            pg_custom_row, textvariable=self._custom_pages_var,
            height=30, placeholder_text="Ví dụ: 1,3,5-8,12",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=8,
        )
        self.custom_pages_entry.bind("<Return>", lambda e: self._commit_custom_pages())
        self.custom_pages_entry.bind("<FocusOut>", lambda e: self._commit_custom_pages())
        self._custom_pages_visible = False
        self._custom_pages_packed = False
        # NOTE: entry packed/unpacked dynamically by _refresh_custom_pages_visibility

        self.page_file_lbl = ctk.CTkLabel(
            card, text="",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=THEME_COLORS["text_muted"],
            anchor="w",
        )
        self.page_file_lbl.grid(row=_section.row, column=0, sticky="ew", padx=14, pady=(0, 2))
        _section.row += 1

        pg_apply_row = _row()
        pg_apply_row.grid_columnconfigure(0, weight=1)
        pg_apply_row.grid_columnconfigure(1, weight=1)
        ctk.CTkButton(
            pg_apply_row, text="📋 Trang cho tất cả", height=28,
            command=self._apply_pages_to_all,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        ).grid(row=0, column=0, sticky="ew", padx=(0, 3))
        ctk.CTkButton(
            pg_apply_row, text="📋 Bản cho tất cả", height=28,
            command=self._apply_default_copies_to_all,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        ).grid(row=0, column=1, sticky="ew", padx=(3, 0))

        # Default copies (compact row)
        def_row = _row()
        ctk.CTkLabel(
            def_row, text="Số bản mặc định:",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=THEME_COLORS["text_muted"],
        ).pack(side="left", padx=(0, 8))
        self.default_copies_entry = ctk.CTkEntry(
            def_row, textvariable=self._default_copies_var,
            width=52, height=28, justify="center",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        )
        self.default_copies_entry.pack(side="left")

        # Detection status + override (per selected file)
        doc_row = _row()
        self.doc_info_lbl = ctk.CTkLabel(
            doc_row, text="",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=THEME_COLORS["text_muted"],
        )
        self.doc_info_lbl.pack(side="left")
        self.btn_duplex_reset = ctk.CTkButton(
            doc_row, text="↩ Về mặc định chung", height=24,
            command=self._reset_file_duplex,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=10),
            corner_radius=6,
        )

        # ══ Section 3: Giấy & kiểu in ══════════════════════════════
        _section("📐  GIẤY & KIỂU IN")

        ctk.CTkLabel(card, text="Khổ giấy:", font=ctk.CTkFont(size=11),
                     text_color=THEME_COLORS["text_muted"]).grid(
            row=_section.row, column=0, sticky="w", padx=14, pady=(0, 0))
        _section.row += 1
        paper_row = _row()
        self._paper_box = ctk.CTkFrame(paper_row, fg_color="transparent")
        self._paper_box.pack(fill="x")
        ctk.CTkComboBox(
            self._paper_box, values=list(PAPER_SIZES.keys()) + [CUSTOM_PAPER_LABEL],
            variable=self._paper_var,
            height=30, state="readonly",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=8,
            command=self._on_paper_change,
        ).pack(fill="x")

        # Custom W×H (hidden unless custom paper chosen)
        self._custom_paper_w = tk.StringVar(value=self._cfg.get("custom_paper_w", "210"))
        self._custom_paper_h = tk.StringVar(value=self._cfg.get("custom_paper_h", "297"))
        self.custom_paper_frame = ctk.CTkFrame(paper_row, fg_color="transparent")
        ctk.CTkLabel(self.custom_paper_frame, text="Rộng (mm):",
                     font=ctk.CTkFont(size=11)).pack(side="left")
        ctk.CTkEntry(self.custom_paper_frame, textvariable=self._custom_paper_w,
                     width=60, height=28).pack(side="left", padx=4)
        ctk.CTkLabel(self.custom_paper_frame, text="× Cao (mm):",
                     font=ctk.CTkFont(size=11)).pack(side="left")
        ctk.CTkEntry(self.custom_paper_frame, textvariable=self._custom_paper_h,
                     width=60, height=28).pack(side="left", padx=4)
        for v in (self._custom_paper_w, self._custom_paper_h):
            v.trace_add("write", lambda *_: self._on_custom_paper_typed())

        ctk.CTkLabel(card, text="Chiều in:", font=ctk.CTkFont(size=11),
                     text_color=THEME_COLORS["text_muted"]).grid(
            row=_section.row, column=0, sticky="w", padx=14, pady=(4, 0))
        _section.row += 1
        orient_row = _row()
        ctk.CTkComboBox(
            orient_row, values=ORIENTATION_OPTIONS, variable=self._orient_var,
            height=30, state="readonly",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=8,
        ).pack(fill="x")

        ctk.CTkLabel(card, text="In 2 mặt:", font=ctk.CTkFont(size=11),
                     text_color=THEME_COLORS["text_muted"]).grid(
            row=_section.row, column=0, sticky="w", padx=14, pady=(4, 0))
        _section.row += 1
        dup_row = _row()
        for text in DUPLEX_MODES:
            ctk.CTkRadioButton(
                dup_row, text=text, variable=self._duplex_var, value=text,
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color=THEME_COLORS["text"],
                radiobutton_width=16, radiobutton_height=16,
            ).pack(anchor="w", pady=1)

        ctk.CTkLabel(card, text="Lề đóng gáy:", font=ctk.CTkFont(size=11),
                     text_color=THEME_COLORS["text_muted"]).grid(
            row=_section.row, column=0, sticky="w", padx=14, pady=(4, 0))
        _section.row += 1
        margin_row = _row()
        self.margin_combo = ctk.CTkComboBox(
            margin_row, values=list(BINDING_MARGIN_OPTIONS.keys()),
            variable=self._binding_margin_var,
            height=30, state="readonly",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=8,
        )
        self.margin_combo.pack(fill="x")

        ctk.CTkLabel(card, text="Máy in dự phòng:", font=ctk.CTkFont(size=11),
                     text_color=THEME_COLORS["text_muted"]).grid(
            row=_section.row, column=0, sticky="w", padx=14, pady=(4, 0))
        _section.row += 1
        failover_row = _row()
        self.failover_combo = ctk.CTkComboBox(
            failover_row, values=["(Không dùng)"],
            variable=self._failover_printer_var,
            height=30, state="readonly",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=8,
        )
        self.failover_combo.pack(fill="x")

        self._build_sec_invoice(card, _section, _row)

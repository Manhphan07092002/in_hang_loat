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


class ActionCardMixin:

    def _build_action_card(self, parent):
        card = ctk.CTkFrame(
            parent, fg_color=THEME_COLORS["card"],
            corner_radius=12, border_width=1,
            border_color=THEME_COLORS["card_border"],
        )
        card.grid(row=1, column=0, sticky="ew")
        card.grid_columnconfigure(0, weight=1)

        # ── Header / Section Title ───────────────────────────────────
        prog_header = ctk.CTkFrame(card, fg_color="transparent")
        prog_header.grid(row=0, column=0, sticky="ew", padx=14, pady=(8, 2))

        ctk.CTkLabel(
            prog_header, text="📊 TIẾN TRÌNH IN",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=THEME_COLORS["text"],
        ).pack(side="left")

        self.status_lbl = ctk.CTkLabel(
            prog_header, text="🟢 Sẵn sàng in",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=THEME_COLORS["text"],
        )
        self.status_lbl.pack(side="right")

        # ── Progress Bars ────────────────────────────────────────────
        prog_wrap = ctk.CTkFrame(card, fg_color="transparent")
        prog_wrap.grid(row=1, column=0, sticky="ew", padx=14, pady=(2, 6))
        prog_wrap.grid_columnconfigure(1, weight=1)

        # File progress
        self.file_progress_title = ctk.CTkLabel(
            prog_wrap, text="Tệp hiện tại:",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=THEME_COLORS["text"],
        )
        self.file_progress_title.grid(row=0, column=0, sticky="w", padx=(0, 6), pady=2)

        self.file_progress = ctk.CTkProgressBar(
            prog_wrap, height=12, corner_radius=6,
            progress_color=THEME_COLORS["primary"][0],
            fg_color=("#CBD5E1", "#334155"),
        )
        self.file_progress.grid(row=0, column=1, sticky="ew", padx=6, pady=2)
        self.file_progress.set(0)

        self.file_progress_lbl = ctk.CTkLabel(
            prog_wrap, text="0/0 trang",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=THEME_COLORS["text"], width=80,
        )
        self.file_progress_lbl.grid(row=0, column=2, sticky="e", pady=2)

        # Total progress
        self.total_progress_title = ctk.CTkLabel(
            prog_wrap, text="Tổng tiến độ:",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=THEME_COLORS["text"],
        )
        self.total_progress_title.grid(row=1, column=0, sticky="w", padx=(0, 6), pady=2)

        self.total_progress = ctk.CTkProgressBar(
            prog_wrap, height=12, corner_radius=6,
            progress_color=THEME_COLORS["success"][0],
            fg_color=("#CBD5E1", "#334155"),
        )
        self.total_progress.grid(row=1, column=1, sticky="ew", padx=6, pady=2)
        self.total_progress.set(0)

        self.total_progress_lbl = ctk.CTkLabel(
            prog_wrap, text="0/0 tệp",
            font=ctk.CTkFont(family="Segoe UI", size=10, weight="bold"),
            text_color=THEME_COLORS["text"], width=80,
        )
        self.total_progress_lbl.grid(row=1, column=2, sticky="e", pady=2)

        # ── Large Action Buttons ─────────────────────────────────────
        btn_bar = ctk.CTkFrame(card, fg_color="transparent")
        btn_bar.grid(row=2, column=0, sticky="ew", padx=14, pady=(2, 6))

        self.btn_start = ctk.CTkButton(
            btn_bar, text="▶️ BẮT ĐẦU IN", height=42,
            command=self.start_print,
            fg_color=THEME_COLORS["success"],
            hover_color=THEME_COLORS["success_hover"],
            text_color="#FFFFFF",
            text_color_disabled="#FFFFFF",
            border_width=1,
            border_color=THEME_COLORS["success_hover"],
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            corner_radius=8,
        )
        self.btn_start.pack(fill="x", pady=(0, 4))

        # Secondary Control Row (Tạm Dừng & Hủy In)
        sub_ctrl_bar = ctk.CTkFrame(card, fg_color="transparent")
        sub_ctrl_bar.grid(row=3, column=0, sticky="ew", padx=14, pady=(0, 4))
        sub_ctrl_bar.grid_columnconfigure(0, weight=1)
        sub_ctrl_bar.grid_columnconfigure(1, weight=1)

        self.btn_pause = ctk.CTkButton(
            sub_ctrl_bar, text="⏸️ Tạm Dừng", height=34,
            command=self.pause_print, state="disabled",
            fg_color=THEME_COLORS["warning"],
            hover_color=THEME_COLORS["warning_hover"],
            text_color="#FFFFFF",
            text_color_disabled="#FFFFFF",
            border_width=1,
            border_color=THEME_COLORS["warning_hover"],
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            corner_radius=8,
        )
        self.btn_pause.grid(row=0, column=0, sticky="ew", padx=(0, 3))

        self.btn_cancel = ctk.CTkButton(
            sub_ctrl_bar, text="⏹️ Hủy In", height=34,
            command=self.cancel_print, state="disabled",
            fg_color=THEME_COLORS["danger"],
            hover_color=THEME_COLORS["danger_hover"],
            text_color="#FFFFFF",
            text_color_disabled="#FFFFFF",
            border_width=1,
            border_color=THEME_COLORS["danger_hover"],
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            corner_radius=8,
        )
        self.btn_cancel.grid(row=0, column=1, sticky="ew", padx=(3, 0))

        # Secondary Action Row (Selective & Retry)
        sub_btn_bar = ctk.CTkFrame(card, fg_color="transparent")
        sub_btn_bar.grid(row=4, column=0, sticky="ew", padx=14, pady=(0, 10))
        sub_btn_bar.grid_columnconfigure(0, weight=1)
        sub_btn_bar.grid_columnconfigure(1, weight=1)

        self.btn_print_selected = ctk.CTkButton(
            sub_btn_bar, text="🎯 Chỉ In Mục Chọn", height=32,
            command=self.start_print_selected,
            fg_color=THEME_COLORS["primary"],
            hover_color=THEME_COLORS["primary_hover"],
            text_color="#FFFFFF",
            text_color_disabled="#FFFFFF",
            border_width=1,
            border_color=THEME_COLORS["primary_hover"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=8,
        )
        self.btn_print_selected.grid(row=0, column=0, sticky="ew", padx=(0, 3))

        self.btn_retry_failed = ctk.CTkButton(
            sub_btn_bar, text="🔄 In Lại Tệp Lỗi / Hủy", height=32,
            command=self.retry_failed_prints,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            text_color_disabled=THEME_COLORS["btn_secondary_text"],
            border_width=1,
            border_color=("#94A3B8", "#475569"),
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=8,
        )
        self.btn_retry_failed.grid(row=0, column=1, sticky="ew", padx=(3, 0))

    # ═══════════════════════════════════════════════════════════════════
    #  QUEUE & FILE MANAGEMENT
    # ═══════════════════════════════════════════════════════════════════

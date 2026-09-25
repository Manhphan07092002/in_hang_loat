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


class PanelsMixin:

    def _build_main_dashboard(self):
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=1, column=0, sticky="nsew", padx=12, pady=10)
        body.grid_rowconfigure(0, weight=1)
        # Responsive: trái co giãn tự do (minsize=0 để không ép cửa sổ),
        # phải cố định ~340px (minsize=320, weight=0 nên không phình khi maximize).
        # MIN_SIZE (860) đảm bảo breakpoint <850 không bao giờ tới được.
        body.grid_columnconfigure(0, weight=1, minsize=0)
        body.grid_columnconfigure(1, weight=0, minsize=320)
        self._body = body

        # ── Column 0: Left Column (Queue Table, Summary & Log Box) ────
        left_col = ctk.CTkFrame(body, fg_color="transparent")
        left_col.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=0)
        left_col.grid_rowconfigure(0, weight=7)  # Queue Card
        left_col.grid_rowconfigure(1, weight=3)  # Log Card
        left_col.grid_columnconfigure(0, weight=1, minsize=0)
        self._left_col = left_col

        self._build_queue_card(left_col)
        self._build_log_card(left_col)

        # ── Column 1: Right Column (Settings, Progress & Control) ────
        right_col = ctk.CTkFrame(body, fg_color="transparent")
        right_col.grid(row=0, column=1, sticky="nsew", padx=(6, 0), pady=0)
        right_col.grid_rowconfigure(0, weight=1)  # Settings Card (Scrollable)
        right_col.grid_rowconfigure(1, weight=0)  # Progress & Actions Card (Docked Bottom)
        right_col.grid_columnconfigure(0, weight=1, minsize=0)
        self._right_col = right_col

        # Printer & Settings Card (at top)
        self._build_settings_card(right_col)

        # Progress & Control Actions Card (at bottom)
        self._build_action_card(right_col)

    # ═══════════════════════════════════════════════════════════════════
    #  LEFT COLUMN: QUEUE CARD (RESPONSIVE TOOLBAR & STRETCHING TABLE)
    # ═══════════════════════════════════════════════════════════════════

    def _build_log_card(self, parent):
        card = ctk.CTkFrame(
            parent, fg_color=THEME_COLORS["card"],
            corner_radius=12, border_width=1,
            border_color=THEME_COLORS["card_border"],
        )
        card.grid(row=1, column=0, sticky="nsew")
        card.grid_rowconfigure(1, weight=1)
        card.grid_columnconfigure(0, weight=1)

        # Header
        top = ctk.CTkFrame(card, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=12, pady=(8, 4))

        ctk.CTkLabel(
            top, text="📋 Nhật Ký Hoạt Động",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=THEME_COLORS["text"],
        ).pack(side="left")

        ctk.CTkButton(
            top, text="🧹 Xóa log", width=70, height=24,
            command=self.clear_log,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=10),
            corner_radius=6,
        ).pack(side="right", padx=(4, 0))

        ctk.CTkButton(
            top, text="💾 Lưu log", width=70, height=24,
            command=self.save_log,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=10),
            corner_radius=6,
        ).pack(side="right", padx=4)

        # Log Text Box
        self.log_box = ctk.CTkTextbox(
            card, height=110,
            font=ctk.CTkFont(family="Consolas", size=10),
            fg_color=THEME_COLORS["card_alt"],
            text_color=THEME_COLORS["text"],
            corner_radius=8,
            state="disabled",
        )
        self.log_box.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 8))

    # ═══════════════════════════════════════════════════════════════════
    #  RIGHT COLUMN: SETTINGS CARD
    # ═══════════════════════════════════════════════════════════════════

    def _log(self, message: str):
        line = f"[{get_timestamp()}]  {message}\n"
        self.log_box.configure(state="normal")
        self.log_box.insert("end", line)
        # Giới hạn log để tránh phình RAM khi in hàng trăm file
        try:
            nlines = int(self.log_box.index("end-1c").split(".")[0])
            if nlines > 2000:
                self.log_box.delete("1.0", f"{nlines - 2000}.0")
        except Exception:
            pass
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        log_info(message)  # mirror sang file %APPDATA%/logs/app.log

    def save_log(self):
        path = filedialog.asksaveasfilename(
            title="Lưu nhật ký",
            defaultextension=".txt",
            filetypes=[("Tệp văn bản (*.txt)", "*.txt"), ("Tất cả tệp (*.*)", "*.*")],
            initialfile=f"nhat_ky_in_{datetime.now():%Y%m%d_%H%M%S}.txt",
        )
        if not path:
            return
        try:
            self.log_box.configure(state="normal")
            content = self.log_box.get("1.0", "end")
            self.log_box.configure(state="disabled")
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            self._log(f"Đã lưu log → {path}")
        except Exception as exc:
            messagebox.showerror("Lỗi", f"Không thể lưu log:\n{exc}")

    def clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    # ── Cleanup on Close ─────────────────────────────────────────────

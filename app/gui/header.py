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
from .paths import get_resource_path
from PIL import Image, ImageTk
import sys


"""Part of PDFBatchPrinterApp (see app/gui/__init__.py)."""


class HeaderMixin:

    def _build_header_bar(self):
        self.header = ctk.CTkFrame(
            self, fg_color=THEME_COLORS["header_bg"],
            corner_radius=0, height=60,
            border_width=1, border_color=THEME_COLORS["card_border"],
        )
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.grid_propagate(False)

        # Left: App Brand & Badge
        self.brand_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        self.brand_frame.pack(side="left", padx=(16, 8), pady=10)

        logo_path = get_resource_path(os.path.join("assets", "logo.png"))
        if os.path.exists(logo_path):
            try:
                logo_pil = Image.open(logo_path)
                self._header_logo_ctk = ctk.CTkImage(
                    light_image=logo_pil,
                    dark_image=logo_pil,
                    size=(34, 34)
                )
                ctk.CTkLabel(
                    self.brand_frame, image=self._header_logo_ctk, text=""
                ).pack(side="left", padx=(0, 8))
            except Exception as e:
                print(f"Error loading header logo: {e}")

        ctk.CTkLabel(
            self.brand_frame, text="PDF BATCH PRINTER PRO",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"),
            text_color=THEME_COLORS["text"],
        ).pack(side="left")

        self.type_badge = ctk.CTkFrame(
            self.brand_frame, fg_color=THEME_COLORS["accent_pill"],
            corner_radius=12, height=24,
        )
        self.type_badge.pack(side="left", padx=(10, 0))

        ctk.CTkLabel(
            self.type_badge, text="PDF • Word • Excel • PowerPoint • Ảnh",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=THEME_COLORS["accent_pill_text"],
        ).pack(padx=10, pady=2)

        # Right: Theme Toggle & Stats Pill
        controls_frame = ctk.CTkFrame(self.header, fg_color="transparent")
        controls_frame.pack(side="right", padx=16, pady=10)

        # Stats Pill
        self.stats_pill = ctk.CTkFrame(
            controls_frame, fg_color=THEME_COLORS["card_alt"],
            corner_radius=12, height=30,
            border_width=1, border_color=THEME_COLORS["card_border"],
        )
        self.stats_pill.pack(side="left", padx=(0, 8))

        self.stats_lbl = ctk.CTkLabel(
            self.stats_pill, text="📦 0 tệp tin  •  📄 0 trang",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=THEME_COLORS["text"],
        )
        self.stats_lbl.pack(padx=10, pady=4)

        # User Guide / Help Button
        self.btn_guide = ctk.CTkButton(
            controls_frame, text="❓ Hướng Dẫn",
            command=self._open_user_guide,
            width=96, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=THEME_COLORS["card_alt"],
            hover_color=THEME_COLORS["primary"][0],
            text_color=THEME_COLORS["text"],
            border_width=1, border_color=THEME_COLORS["card_border"],
            corner_radius=8,
        )
        self.btn_guide.pack(side="left", padx=(0, 10))

        # Update Checker Button (badge dot appears when a new version exists)
        self.btn_update = ctk.CTkButton(
            controls_frame, text="🔄 Cập nhật",
            command=lambda: self.check_updates(manual=True),
            width=96, height=28,
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            fg_color=THEME_COLORS["card_alt"],
            hover_color=THEME_COLORS["primary"][0],
            text_color=THEME_COLORS["text"],
            border_width=1, border_color=THEME_COLORS["card_border"],
            corner_radius=8,
        )
        self.btn_update.pack(side="left", padx=(0, 10))

        # Dark/Light Mode Switch
        self.theme_switch = ctk.CTkSwitch(
            controls_frame, text="🌙 Chế độ Tối",
            command=self._toggle_theme,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=THEME_COLORS["text"],
            progress_color=THEME_COLORS["primary"][0],
        )
        self.theme_switch.pack(side="left")

        self.header.bind("<Configure>", self._on_header_resize)

    def _open_user_guide(self):
        """Open the user manual HTML guide in the default web browser."""
        import webbrowser
        possible_paths = [
            get_resource_path("HUONG_DAN_SU_DUNG.html"),
            get_resource_path(os.path.join("assets", "HUONG_DAN_SU_DUNG.html")),
            os.path.join(get_app_data_dir(), "HUONG_DAN_SU_DUNG.html"),
            os.path.join(get_app_data_dir(), "assets", "HUONG_DAN_SU_DUNG.html"),
        ]
        guide_path = None
        for p in possible_paths:
            if os.path.exists(p):
                guide_path = p
                break

        if guide_path:
            try:
                os.startfile(guide_path)
            except Exception:
                webbrowser.open(f"file:///{os.path.abspath(guide_path).replace(os.sep, '/')}")
        else:
            messagebox.showwarning("Hướng Dẫn Sử Dụng", "Không tìm thấy tệp hướng dẫn HUONG_DAN_SU_DUNG.html")

    # ── Self-update (GitHub Releases) ────────────────────────────────

    def _on_header_resize(self, event):
        """Hide the type badge below 960px so header never clips.

        Guarded by state: layout changes re-fire <Configure>, so only act
        on transitions (no flicker loop). MIN_SIZE width is 860, therefore
        deeper collapsing is unnecessary — toolbar/table/scrollbars cover it.
        """
        try:
            narrow = event.width < 960
            if narrow == getattr(self, "_header_narrow", None):
                return
            self._header_narrow = narrow
            if narrow:
                if self.type_badge.winfo_ismapped():
                    self.type_badge.pack_forget()
            else:
                if not self.type_badge.winfo_ismapped():
                    self.type_badge.pack(side="left", padx=(10, 0))
        except Exception:
            pass

    def _toggle_theme(self):
        if self.theme_switch.get() == 1:
            ctk.set_appearance_mode("Dark")
            label = "☀️ Chế độ Sáng"
        else:
            ctk.set_appearance_mode("Light")
            label = "🌙 Chế độ Tối"
        try:
            self.theme_switch.configure(text=label)
        except Exception:
            pass

        self.configure(fg_color=THEME_COLORS["bg"])
        self._update_treeview_theme()
        # Repaint các vùng dùng màu cứng theo theme
        try:
            is_dark = ctk.get_appearance_mode() == "Dark"
            if hasattr(self, "_table_container"):
                self._table_container.configure(bg="#1E293B" if is_dark else "#FFFFFF")
        except Exception:
            pass

    # ═══════════════════════════════════════════════════════════════════
    #  MAIN 2-COLUMN DASHBOARD (RESPONSIVE)
    # ═══════════════════════════════════════════════════════════════════

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


class QueueViewMixin:

    def _build_queue_card(self, parent):
        card = ctk.CTkFrame(
            parent, fg_color=THEME_COLORS["card"],
            corner_radius=12, border_width=1,
            border_color=THEME_COLORS["card_border"],
        )
        card.grid(row=0, column=0, sticky="nsew", pady=(0, 6))
        card.grid_rowconfigure(1, weight=1)
        card.grid_columnconfigure(0, weight=1)

        # ── Responsive Toolbar: 1 khung flow duy nhất, nút tự xuống
        #  dòng theo chiều rộng (grid, KHÔNG pack跨-frame vì Tk cấm) ──
        toolbar = ctk.CTkFrame(card, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 6))
        toolbar.grid_columnconfigure(0, weight=1)
        self._toolbar = toolbar

        ctk.CTkLabel(
            toolbar, text="📁 Hàng Đợi",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=THEME_COLORS["text"],
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))

        flow = ctk.CTkFrame(toolbar, fg_color="transparent")
        flow.grid(row=1, column=0, sticky="ew")
        self._toolbar_flow = flow

        self._toolbar_row1_btns = []
        self._toolbar_row2_btns = []
        self._toolbar_buttons = []  # thứ tự đọc ổn định cho reflow
        for text, cmd, fg, hv, txt_col in [
            ("➕ Thêm File", self.add_files, THEME_COLORS["primary"], THEME_COLORS["primary_hover"], "#FFFFFF"),
            ("📁 Thư Mục", self.add_folder, THEME_COLORS["primary"], THEME_COLORS["primary_hover"], "#FFFFFF"),
            ("✏️ Sửa Bản", self._edit_selected_copies, THEME_COLORS["btn_secondary"], THEME_COLORS["btn_secondary_hover"], THEME_COLORS["btn_secondary_text"]),
            ("▲", self.move_selected_up, THEME_COLORS["btn_secondary"], THEME_COLORS["btn_secondary_hover"], THEME_COLORS["btn_secondary_text"]),
            ("▼", self.move_selected_down, THEME_COLORS["btn_secondary"], THEME_COLORS["btn_secondary_hover"], THEME_COLORS["btn_secondary_text"]),
            ("💾 Lưu DS", self.save_queue_session, THEME_COLORS["btn_secondary"], THEME_COLORS["btn_secondary_hover"], THEME_COLORS["btn_secondary_text"]),
            ("📂 Mở DS", self.load_queue_session, THEME_COLORS["btn_secondary"], THEME_COLORS["btn_secondary_hover"], THEME_COLORS["btn_secondary_text"]),
            ("🗑️ Xóa Chọn", self.remove_selected, THEME_COLORS["btn_secondary"], THEME_COLORS["btn_secondary_hover"], THEME_COLORS["btn_secondary_text"]),
            ("🧹 Xóa Hết", self.remove_all, THEME_COLORS["danger"], THEME_COLORS["danger_hover"], "#FFFFFF"),
        ]:
            b = ctk.CTkButton(
                flow, text=text, command=cmd,
                height=28,
                fg_color=fg, hover_color=hv,
                text_color=txt_col,
                text_color_disabled=txt_col,
                font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
                corner_radius=6,
            )
            self._toolbar_buttons.append(b)
        self._toolbar_row1_btns = list(self._toolbar_buttons[:5])
        self._toolbar_row2_btns = list(self._toolbar_buttons[5:])

        self._subfolder_check = ctk.CTkCheckBox(
            flow, text="Quét thư mục con",
            variable=self._include_subfolders,
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=THEME_COLORS["text"],
            checkbox_width=16, checkbox_height=16,
            corner_radius=4,
        )

        self._reflow_job = None
        self._reflowing = False
        toolbar.bind("<Configure>", self._on_toolbar_resize)
        self._reflow_toolbar()

        # ── Modern Treeview Table ────────────────────────────────────
        table_container = tk.Frame(card, bg="#FFFFFF", bd=0, highlightthickness=0)
        self._table_container = table_container  # giữ ref để repaint khi đổi theme
        table_container.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 4))
        table_container.grid_rowconfigure(0, weight=1)
        table_container.grid_columnconfigure(0, weight=1)

        cols = ("stt", "filename", "filetype", "pages", "size", "pagesel", "copies", "status", "doctype", "duplex")
        self.tree = ttk.Treeview(
            table_container, columns=cols, show="headings",
            selectmode="extended", style="Modern.Treeview",
        )
        
        # Interactive Sortable Headings
        headings_map = [
            ("stt", "#"),
            ("filename", "Tên Tệp Tin ⇕"),
            ("filetype", "Định Dạng ⇕"),
            ("pages", "Số Trang ⇕"),
            ("size", "Kích Thước ⇕"),
            ("pagesel", "Trang In ⇕"),
            ("copies", "Số Bản (✎) ⇕"),
            ("status", "Trạng Thái ⇕"),
            ("doctype", "Loại ⇕"),
            ("duplex", "2 Mặt ⇕"),
        ]
        for col_id, h_text in headings_map:
            self.tree.heading(
                col_id, text=h_text,
                command=lambda c=col_id: self._sort_by_column(c),
            )

        self.tree.column("stt", width=36, minwidth=32, anchor="center")
        self.tree.column("filename", width=200, minwidth=100)
        self.tree.column("filetype", width=85, minwidth=75, anchor="center")
        self.tree.column("pages", width=70, minwidth=60, anchor="center")
        self.tree.column("size", width=85, minwidth=75, anchor="center")
        self.tree.column("pagesel", width=95, minwidth=85, anchor="center")
        self.tree.column("copies", width=90, minwidth=70, anchor="center")
        self.tree.column("status", width=105, minwidth=95, anchor="center")
        self.tree.column("doctype", width=90, minwidth=75, anchor="center")
        self.tree.column("duplex", width=85, minwidth=70, anchor="center")

        vsb = ttk.Scrollbar(table_container, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # Horizontal scrollbar: bảng tự cuộn ngang khi cửa sổ hẹp,
        # không bao giờ ép vỡ layout hay tạo scrollbar ngang toàn app.
        hsb = ttk.Scrollbar(table_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky="ew")
        self._tree_hsb = hsb

        # Bind container configure event for dynamic filename column auto-stretching
        table_container.bind("<Configure>", self._on_table_resize)

        self.tree.bind("<<TreeviewSelect>>", self._on_file_select)
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Button-3>", self._on_tree_right_click)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())
        self.tree.bind("<Control-a>", self._select_all_rows)
        self.tree.bind("<Control-Up>", lambda e: self.move_selected_up())
        self.tree.bind("<Control-Down>", lambda e: self.move_selected_down())
        self.tree.bind("<Alt-Up>", lambda e: self.move_selected_up())
        self.tree.bind("<Alt-Down>", lambda e: self.move_selected_down())
        self.tree.bind("<Key>", self._on_tree_keypress)
        self.tree.bind("<Motion>", self._on_tree_hover)
        self.tree.bind("<Leave>", lambda e: self._hide_tree_tooltip())
        self._tree_tip = None
        self._tree_tip_job = None

        self._update_treeview_theme()

        # ── Queue Summary Information Bar (Thanh Tổng Hợp) ──────────
        summary_frame = ctk.CTkFrame(
            card, fg_color=THEME_COLORS["card_alt"],
            corner_radius=8, height=32,
        )
        summary_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        summary_frame.pack_propagate(False)

        self.queue_summary_lbl = ctk.CTkLabel(
            summary_frame,
            text="📦 Tổng tài liệu: 0  •  📑 Tổng số bản: 0  •  📄 Tổng số trang in: 0 trang",
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            text_color=THEME_COLORS["text"],
            anchor="w",
        )
        self.queue_summary_lbl.pack(side="left", padx=12, pady=4)

    def _on_table_resize(self, event):
        """Dynamically expand the filename column to 100% of available space without clipping."""
        try:
            total_w = event.width
            # Sum of fixed columns (stt:36, filetype:85, pages:70, size:85, pagesel:95, copies:90, status:105, doctype:90, duplex:85) + scrollbar margin
            fixed_w = 36 + 85 + 70 + 85 + 95 + 90 + 105 + 90 + 85 + 24
            rem_w = max(100, total_w - fixed_w)
            self.tree.column("filename", width=rem_w)
        except Exception:
            pass

    def _update_treeview_theme(self):
        """Configure ttk Treeview styles for clean modern appearance."""
        is_dark = ctk.get_appearance_mode() == "Dark"
        bg_color = "#1E293B" if is_dark else "#FFFFFF"
        fg_color = "#F8FAFC" if is_dark else "#0F172A"
        heading_bg = "#0F172A" if is_dark else "#F1F5F9"
        heading_fg = "#94A3B8" if is_dark else "#475569"
        select_bg = "#2563EB" if is_dark else "#DBEAFE"
        select_fg = "#FFFFFF" if is_dark else "#1E40AF"

        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Modern.Treeview",
            background=bg_color,
            foreground=fg_color,
            rowheight=28,
            fieldbackground=bg_color,
            font=("Segoe UI", 10),
            borderwidth=0,
        )
        style.configure(
            "Modern.Treeview.Heading",
            background=heading_bg,
            foreground=heading_fg,
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padding=(4, 3),
        )
        style.map(
            "Modern.Treeview",
            background=[("selected", select_bg)],
            foreground=[("selected", select_fg)],
        )

    # ═══════════════════════════════════════════════════════════════════
    #  LEFT COLUMN: LOG CARD (RESPONSIVE)
    # ═══════════════════════════════════════════════════════════════════

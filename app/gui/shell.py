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


class AppShell(ctk.CTk):

    def __init__(self):
        super().__init__()

        # ── Load Persistent Config ───────────────────────────────────
        self._cfg = self._load_config()
        saved_theme = self._cfg.get("theme", "Light")
        ctk.set_appearance_mode(saved_theme)

        # ── Window Setup & App Icon ──────────────────────────────────
        self.title(WINDOW_TITLE)
        self.minsize(*MIN_SIZE)
        self.configure(fg_color=THEME_COLORS["bg"])
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._set_app_icon()

        # ── State ────────────────────────────────────────────────────
        self._is_alive = True
        import queue
        self._async_queue = queue.Queue()
        self.pdf_manager = PDFManager()
        self.file_list: list[FileInfo] = []
        self.print_worker: Optional[MultiPrinterCoordinator] = None
        self._is_printing = False

        # Restored Tk variables from config
        self._copies_var = tk.IntVar(value=self._cfg.get("copies", 1))
        self._default_copies_var = tk.IntVar(value=1)
        self._selected_filename_var = tk.StringVar(value="(Chưa chọn tệp)")
        self._page_range_var = tk.StringVar(value=self._cfg.get("page_range", PAGE_RANGE_ALL))
        self._custom_pages_var = tk.StringVar(value="")
        _saved_paper = self._cfg.get("paper", "A4")
        try:
            resolve_paper(_saved_paper)
            _paper_init = _saved_paper
        except Exception:
            _paper_init = "A4"
        self._paper_var = tk.StringVar(value=_paper_init)
        self._orient_var = tk.StringVar(value=self._cfg.get("orientation", ORIENT_AUTO))
        self._duplex_var = tk.StringVar(value=self._cfg.get("duplex", "1 mặt"))
        self._saved_printer = self._cfg.get("printer", "")
        self._printer_var = tk.StringVar(value="")
        self._fit_to_page = tk.BooleanVar(value=self._cfg.get("fit_to_page", True))
        self._include_subfolders = tk.BooleanVar(value=self._cfg.get("include_subfolders", False))
        self._theme_switch_var = tk.StringVar(value=saved_theme)

        # Smart Printing Variables
        self._remove_blanks_var = tk.BooleanVar(value=self._cfg.get("remove_blanks", False))
        self._separator_sheet_var = tk.BooleanVar(value=self._cfg.get("separator_sheet", False))
        self._reverse_order_var = tk.BooleanVar(value=self._cfg.get("reverse_order", False))
        self._binding_margin_var = tk.StringVar(value=self._cfg.get("binding_margin", "0 mm (Chuẩn)"))
        self._saved_failover = self._cfg.get("failover_printer", "(Không dùng)")
        _inv = self._cfg.get("invoice_mode", INVOICE_AUTO)
        if _inv not in INVOICE_MODES:
            _inv = INVOICE_AUTO
        self._invoice_mode_var = tk.StringVar(value=_inv)
        self._failover_printer_var = tk.StringVar(value=self._saved_failover)
        self._parallel_printers: list[str] = []
        self._available_printers: list[str] = []

        self._edit_entry: Optional[tk.Entry] = None
        self._sort_reverse: dict[str, bool] = {}
        self._syncing_page_ui: bool = False

        # ── Build Layout ─────────────────────────────────────────────
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_header_bar()
        self._build_main_dashboard()

        # Sync Theme Switch state
        if saved_theme == "Dark":
            self.theme_switch.select()
            self.theme_switch.configure(text="☀️ Chế độ Sáng")

        # ── Drag & Drop Integration ──────────────────────────────────
        self._init_drag_and_drop()

        # ── Show and Position Window on Desktop ──────────────────────
        self.show_window()

        # ── Global Key Bindings ──────────────────────────────────────
        self.bind("<F1>", lambda e: self._open_user_guide())

        # ── Start Async Queue Polling ────────────────────────────────
        self._poll_async_queue()

        # ── Load Printers (Async) ────────────────────────────────────
        self.load_printers()
        self._update_stats_pill()

        # ── Silent self-update check (once/day, background) ──────────
        import threading as _th
        _th.Thread(target=self.auto_check_updates, daemon=True).start()

    def show_window(self):
        """Calculates proper position, centers on screen, and forces the window visible in foreground."""
        try:
            self.update_idletasks()
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            w = min(DEFAULT_SIZE[0], max(MIN_SIZE[0], sw - 60))
            h = min(DEFAULT_SIZE[1], max(MIN_SIZE[1], sh - 80))
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 2)
            self.geometry(f"{w}x{h}+{x}+{y}")
        except Exception:
            self.geometry(f"{DEFAULT_SIZE[0]}x{DEFAULT_SIZE[1]}")

        self.deiconify()
        self.state("normal")
        self.lift()
        self.attributes("-topmost", True)
        self.focus_force()
        self.after(200, self._restore_topmost)
        self.after(500, self._safety_deiconify)

    def _restore_topmost(self):
        try:
            if getattr(self, "_is_alive", False):
                self.attributes("-topmost", False)
        except Exception:
            pass

    def _safety_deiconify(self):
        """Second-pass check to ensure window is deiconified even after Windows DWM theme hooks."""
        try:
            if getattr(self, "_is_alive", False):
                if self.state() == "withdrawn":
                    self.deiconify()
                self.lift()
        except Exception:
            pass

    # ── Config Persistence ───────────────────────────────────────────

    def _poll_async_queue(self):
        """Safely process background thread messages on the main GUI thread."""
        if not getattr(self, "_is_alive", False):
            return
        try:
            while not self._async_queue.empty():
                func, args = self._async_queue.get_nowait()
                try:
                    func(*args)
                except Exception as exc:
                    print(f"Async queue error: {exc}")
        except Exception:
            pass
        finally:
            if getattr(self, "_is_alive", False):
                try:
                    self.after(50, self._poll_async_queue)
                except Exception:
                    pass

    def _set_app_icon(self):
        """Set custom icon for window title bar and Windows taskbar."""
        try:
            import ctypes
            app_id = "PDFBatchPrinterPro.App.1.0"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        except Exception:
            pass

        ico_path = get_resource_path(os.path.join("assets", "app.ico"))
        if os.path.exists(ico_path):
            try:
                self.iconbitmap(default=ico_path)
            except Exception:
                try:
                    self.iconbitmap(ico_path)
                except Exception:
                    pass

        logo_path = get_resource_path(os.path.join("assets", "logo.png"))
        if os.path.exists(logo_path):
            try:
                icon_img = Image.open(logo_path)
                self._app_icon_photo = ImageTk.PhotoImage(icon_img)
                self.iconphoto(True, self._app_icon_photo)
            except Exception:
                pass

    # ═══════════════════════════════════════════════════════════════════
    # ═══════════════════════════════════════════════════════════════════
    #  HEADER BAR (RESPONSIVE)
    # ═══════════════════════════════════════════════════════════════════

    def _on_close(self):
        try:
            self._save_config()
        except Exception:
            pass
        if self._is_printing:
            if not messagebox.askyesno(
                "Xác nhận thoát",
                "Quá trình in đang diễn ra! Bạn có chắc muốn dừng và thoát ứng dụng?",
            ):
                return
            if self.print_worker:
                self.print_worker.cancel()
        self._is_alive = False
        cleanup_temp()
        try:
            self.quit()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

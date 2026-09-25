"""
Modern GUI for PDF Batch Printer Pro — built with CustomTkinter.

Modern Dashboard Layout:
  ┌─ Header Bar ──────────────────────────────────────────────────────────────────┐
  │  🖨️ PDF BATCH PRINTER PRO   [11 tệp • 238 trang]   [🌙 Giao diện Tối / Sáng]  │
  ├────────────────────────────────────────┬──────────────────────────────────────┤
  │ 📂 HÀNG ĐỢI IN (QUEUE)                 │ 👁️ XEM TRƯỚC (PREVIEW)               │
  │  • Toolbar (+File, +Thư mục, Xóa...)   │  • Document canvas + Page navigation │
  │  • Modern Table with Status Badges     ├──────────────────────────────────────┤
  │  • Double-click to edit copies         │ ⚙️ THIẾT LẬP MÁY IN & THÔNG SỐ       │
  ├────────────────────────────────────────┤  • Máy in + Status dot + Refresh     │
  │ 📋 NHẬT KÝ IN                          │  • Số bản, Trang in, Khổ giấy, Chiều │
  │  • Real-time log + Lưu / Xóa           │  • Chế độ in 2 mặt + Vừa trang       │
  │                                        ├──────────────────────────────────────┤
  │                                        │ 🚀 TIẾN TRÌNH & ĐIỀU KHIỂN           │
  │                                        │  • Dual Progress Bars                │
  │                                        │  • [▶ BẮT ĐẦU IN] [⏸] [⏹]            │
  └────────────────────────────────────────┴──────────────────────────────────────┘
"""

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

def get_resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# NOTE: get_app_data_dir / CONFIG_FILE are imported from app.config_store
# (single source of truth, %APPDATA%-based) and re-exported for compatibility.

# Set theme globally before widget creation
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


class PDFBatchPrinterApp(ctk.CTk):
    """Modern Desktop GUI for PDF Batch Printer Pro."""

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

    def _load_config(self) -> dict:
        return load_config()

    def _save_config(self):
        try:
            printer_name = self._printer_var.get()
            if printer_name.startswith("("):
                printer_name = ""
            cfg = {
                "printer": printer_name,
                "duplex": self._duplex_var.get(),
                "paper": self._resolve_gui_paper(),
                "custom_paper_w": self._custom_paper_w.get() if hasattr(self, "_custom_paper_w") else "210",
                "custom_paper_h": self._custom_paper_h.get() if hasattr(self, "_custom_paper_h") else "297",
                "orientation": self._orient_var.get(),
                "copies": self._copies_var.get(),
                "fit_to_page": self._fit_to_page.get(),
                "include_subfolders": self._include_subfolders.get(),
                "theme": "Dark" if self.theme_switch.get() == 1 else "Light",
                "remove_blanks": self._remove_blanks_var.get(),
                "separator_sheet": self._separator_sheet_var.get(),
                "reverse_order": self._reverse_order_var.get(),
                "binding_margin": self._binding_margin_var.get(),
                "failover_printer": self._failover_printer_var.get(),
                "invoice_mode": self._invoice_mode_var.get() if hasattr(self, "_invoice_mode_var") else INVOICE_AUTO,
                "update_last_check": self._cfg.get("update_last_check", ""),
                "update_skip_version": self._cfg.get("update_skip_version", ""),
            }
            save_config(cfg)
        except Exception as exc:
            print(f"Error saving config: {exc}")

    # ── Drag & Drop Handling ─────────────────────────────────────────

    def _init_drag_and_drop(self):
        try:
            import windnd
            windnd.hook_dropfiles(self, func=self._on_files_dropped)
        except Exception as exc:
            print(f"Could not initialize windnd drag-and-drop: {exc}")

    def _on_files_dropped(self, files):
        if not getattr(self, "_is_alive", False):
            return
        decoded_paths = []
        for item in files:
            if isinstance(item, bytes):
                try:
                    p = item.decode("utf-8")
                except UnicodeDecodeError:
                    try:
                        p = item.decode("mbcs")
                    except Exception:
                        p = str(item)
            else:
                p = str(item)
            decoded_paths.append(p)

        all_paths_to_add = []
        recursive = self._include_subfolders.get()

        for p in decoded_paths:
            if not os.path.exists(p):
                continue
            if os.path.isdir(p):
                if recursive:
                    for root, _dirs, fnames in os.walk(p):
                        for fn in fnames:
                            ext = os.path.splitext(fn)[1].lower()
                            if ext in SUPPORTED_EXTENSIONS:
                                all_paths_to_add.append(os.path.join(root, fn))
                else:
                    for fn in os.listdir(p):
                        fp = os.path.join(p, fn)
                        if os.path.isfile(fp):
                            ext = os.path.splitext(fn)[1].lower()
                            if ext in SUPPORTED_EXTENSIONS:
                                all_paths_to_add.append(fp)
            elif os.path.isfile(p):
                ext = os.path.splitext(p)[1].lower()
                if ext in SUPPORTED_EXTENSIONS:
                    all_paths_to_add.append(p)

        if all_paths_to_add:
            self._async_queue.put((self._add_to_queue, (all_paths_to_add,)))
        else:
            self._async_queue.put((self._log, ("⚠️ Không tìm thấy tệp tin được hỗ trợ trong các mục vừa kéo thả",)))

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

    def auto_check_updates(self):
        """Silent background check once per day (skipped-version aware)."""
        try:
            last = self._cfg.get("update_last_check", "")
            if not app_updater.should_auto_check(last):
                return
            skip = (self._cfg.get("update_skip_version", "") or "").strip()
            info = app_updater.check_for_updates()
            self._cfg["update_last_check"] = datetime.now().isoformat(timespec="seconds")
            self._save_config()
            if info.get("has_update") and info.get("latest") != skip:
                self._async_queue.put((self._prompt_update, (info, False)))
            elif info.get("has_update"):
                self._mark_update_available(info)
        except Exception as exc:
            print(f"Auto update check failed: {exc}")

    def check_updates(self, manual: bool = False):
        """Manual check from header button (always reports result)."""
        self.btn_update.configure(text="⏳ Đang kiểm tra...", state="disabled")

        def _worker():
            try:
                info = app_updater.check_for_updates()
                self._cfg["update_last_check"] = datetime.now().isoformat(timespec="seconds")
                self._save_config()
                self._async_queue.put((self._on_update_checked, (info, manual)))
            except Exception as exc:
                self._async_queue.put((self._on_update_checked,
                                       ({"has_update": False, "error": str(exc)}, manual)))

        import threading
        threading.Thread(target=_worker, daemon=True).start()

    def _on_update_checked(self, info: dict, manual: bool):
        try:
            self.btn_update.configure(text="🔄 Cập nhật", state="normal")
        except Exception:
            pass
        skip = (self._cfg.get("update_skip_version", "") or "").strip()
        if info.get("has_update") and info.get("latest") != skip:
            self._mark_update_available(info)
            self._prompt_update(info, manual)
        elif manual:
            if info.get("error"):
                messagebox.showwarning("Cập nhật", f"Không kiểm tra được:\n{info['error']}")
            else:
                messagebox.showinfo("Cập nhật", f"Bạn đang dùng bản mới nhất (v{info.get('current', '')}).")

    def _mark_update_available(self, info: dict):
        try:
            self.btn_update.configure(text=f"🔄 Mới: v{info.get('latest', '')}!")
        except Exception:
            pass

    def _prompt_update(self, info: dict, manual: bool):
        if not getattr(self, "_is_alive", False):
            return
        self._mark_update_available(info)
        notes = (info.get("notes") or "").strip()
        if len(notes) > 800:
            notes = notes[:800] + "\n..."
        msg = (f"Có bản mới v{info.get('latest')} (bạn đang dùng v{info.get('current')}).\n\n"
               f"{notes}\n\nTải về ngay?")
        ans = messagebox.askyesnocancel("Cập nhật phần mềm", msg + "\n\n(Yes=Tải về • No=Để sau • Cancel=Bỏ qua bản này)")
        if ans is True:
            self._download_update(info)
        elif ans is None:
            self._cfg["update_skip_version"] = info.get("latest", "")
            self._save_config()
            try:
                self.btn_update.configure(text="🔄 Cập nhật")
            except Exception:
                pass

    def _download_update(self, info: dict):
        url = info.get("url", "")
        if not url or not url.startswith("http"):
            messagebox.showinfo("Cập nhật", "Bản này không có file đính kèm, mở trang release để tải tay:\n" + url)
            import webbrowser
            try:
                webbrowser.open(url)
            except Exception:
                pass
            return

        dlg = ctk.CTkToplevel(self)
        dlg.title("Đang tải bản cập nhật...")
        dlg.geometry("400x140")
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)
        ctk.CTkLabel(dlg, text=f"⏳ Đang tải {info.get('asset_name') or 'bản cập nhật'}...",
                     font=ctk.CTkFont(size=12, weight="bold")).pack(padx=16, pady=(16, 8))
        bar = ctk.CTkProgressBar(dlg, width=340)
        bar.pack(padx=16, pady=4)
        bar.set(0)
        lbl = ctk.CTkLabel(dlg, text="0%")
        lbl.pack(padx=16, pady=(0, 10))

        def _progress(done: int, total: int):
            frac = (done / total) if total else 0
            self._async_queue.put((lambda: (bar.set(frac), lbl.configure(
                text=f"{int(frac*100)}% ({done//1024//1024} MB)" if total else f"{done//1024//1024} MB")), ()))

        def _worker():
            try:
                path = app_updater.download_asset(url, progress_cb=_progress)
                self._async_queue.put((lambda: self._on_update_downloaded(dlg, path), ()))
            except Exception as exc:
                self._async_queue.put((lambda: (
                    dlg.destroy(),
                    messagebox.showerror("Cập nhật", f"Tải thất bại:\n{exc}")), ()))

        import threading
        threading.Thread(target=_worker, daemon=True).start()

    def _on_update_downloaded(self, dlg, path: str):
        try:
            dlg.destroy()
        except Exception:
            pass
        self._log(f"Đã tải bản cập nhật → {path}")
        if messagebox.askyesno("Cập nhật", f"Đã tải xong:\n{path}\n\nChạy cài đặt ngay? (App sẽ thoát)"):
            try:
                import subprocess
                subprocess.Popen([path], close_fds=True)
            except Exception:
                try:
                    os.startfile(path)
                except Exception as exc:
                    messagebox.showerror("Cập nhật", f"Không chạy được file:\n{exc}")
                    return
            self._on_close()

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

    def _on_toolbar_resize(self, _event=None):
        """Debounce toolbar reflow on width change (skip if width unchanged)."""
        try:
            w = self._toolbar.winfo_width()
            if w == getattr(self, "_toolbar_last_w", None):
                return
            self._toolbar_last_w = w
            if self._reflow_job is not None:
                self.after_cancel(self._reflow_job)
        except Exception:
            pass
        try:
            self._reflow_job = self.after(120, self._reflow_toolbar)
        except Exception:
            pass

    def _reflow_toolbar(self):
        """Xếp 9 nút + checkbox vào các dòng grid vừa khung (tham lam theo
        chiều rộng thực). Không nút nào bị cắt/chồng; thứ tự đọc giữ nguyên.
        """
        self._reflow_job = None
        if getattr(self, "_reflowing", False):
            return
        if not getattr(self, "_is_alive", False):
            return
        self._reflowing = True
        try:
            flow = self._toolbar_flow
            try:
                avail = max(220, flow.winfo_width() or 0)
            except Exception:
                return
            if avail <= 0:
                return

            def _w(w):
                try:
                    return w.winfo_reqwidth() + 5
                except Exception:
                    return 80

            rows: list[list] = []
            cur: list = []
            curw = 0
            for b in self._toolbar_buttons:
                bw = _w(b)
                if cur and curw + bw > avail:
                    rows.append(cur)
                    cur, curw = [], 0
                cur.append(b)
                curw += bw
            if cur:
                rows.append(cur)

            for b in self._toolbar_buttons:
                try:
                    b.grid_forget()
                except Exception:
                    pass
            try:
                self._subfolder_check.grid_forget()
            except Exception:
                pass
            for r, rowbtns in enumerate(rows):
                for c, b in enumerate(rowbtns):
                    b.grid(row=r, column=c, padx=2, pady=2, sticky="w")
            # Checkbox luôn cuối dòng cuối, dính phải
            flow.grid_columnconfigure(len(rows[-1]) if rows else 0, weight=1)
            self._subfolder_check.grid(row=len(rows) - 1, column=len(rows[-1]),
                                       padx=(8, 0), pady=2, sticky="e")
            # Cập nhật membership để test/inspect biết nút đang dòng nào
            self._toolbar_row1_btns = list(rows[0]) if rows else []
            self._toolbar_row2_btns = [b for r in rows[1:] for b in r]
        finally:
            self._reflowing = False

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

    def _on_tree_hover(self, event):
        """Show full filename tooltip when hovering a truncated filename cell."""
        try:
            if self._tree_tip_job is not None:
                self.after_cancel(self._tree_tip_job)
                self._tree_tip_job = None
        except Exception:
            pass
        self._hide_tree_tooltip()

        def _show():
            try:
                if self.tree.identify("region", event.x, event.y) != "cell":
                    return
                if self.tree.identify_column(event.x) != "#2":  # filename column
                    return
                item = self.tree.identify_row(event.y)
                if not item:
                    return
                idx = int(self.tree.item(item, "values")[0]) - 1
                if not (0 <= idx < len(self.file_list)):
                    return
                full = self.file_list[idx].filename
                # Chỉ hiện khi text thực sự bị cắt (đo rộng hơn cột)
                from tkinter import font as tkfont
                w = tkfont.Font(family="Segoe UI", size=10).measure(full)
                if w <= self.tree.column("filename", option="width") - 6:
                    return
                tip = tk.Toplevel(self)
                tip.wm_overrideredirect(True)
                tip.wm_geometry(f"+{event.x_root + 14}+{event.y_root + 12}")
                lbl = tk.Label(tip, text=full, bg="#0F172A", fg="#F8FAFC",
                               font=("Segoe UI", 10), padx=8, pady=4,
                               wraplength=420, justify="left")
                lbl.pack()
                self._tree_tip = tip
            except Exception:
                pass

        try:
            self._tree_tip_job = self.after(450, _show)
        except Exception:
            pass

    def _hide_tree_tooltip(self):
        try:
            if self._tree_tip is not None:
                self._tree_tip.destroy()
        except Exception:
            pass
        self._tree_tip = None

    def _select_all_rows(self, _e=None):
        self.tree.selection_set(self.tree.get_children())
        return "break"

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

    def _on_invoice_mode_change(self):
        self._save_config()
        mode = self._invoice_mode_var.get()
        if mode == INVOICE_OFF:
            self._log("Đã tắt nhận diện hóa đơn điện tử")
        else:
            self._log(f"Nhận diện hóa đơn: {mode} — phân tích lại hàng đợi...")
            self._reanalyze_queue_invoice()

    def _reanalyze_queue_invoice(self):
        """Chạy lại nhận diện cho file chưa phân tích (nền, không chặn UI)."""
        import threading

        def _worker():
            changed = []

            def _done():
                if not getattr(self, "_is_alive", False):
                    return
                if changed:
                    self._refresh_tree(preserve_selection=True)
                    self._sync_page_ui_from_selection()
                    for name in changed:
                        self._log(f"🧾 Phát hiện hóa đơn điện tử: {name}")
                    self._log(f"🧾 Nhận diện xong {len(changed)} hóa đơn trong hàng đợi")

            for f in self.file_list:
                if not f.invoice_analyzed:
                    self._analyze_invoice(f)
                    if f.invoice_detected:
                        changed.append(f.filename)
            self._async_queue.put((_done, ()))

        threading.Thread(target=_worker, daemon=True).start()

    def _analyze_invoice(self, finfo) -> bool:
        """Phân tích 1 file, điền document_type/invoice_*. Không bao giờ raise."""
        if getattr(finfo, "invoice_analyzed", False):
            return bool(getattr(finfo, "invoice_detected", False))
        try:
            if getattr(finfo, "is_converted", False) and not getattr(finfo, "is_converted_ready", False):
                return False  # Office chưa convert — hẹn lúc chuẩn bị in
            pdf = getattr(finfo, "pdf_path", "") or ""
            if not pdf.lower().endswith(".pdf") or not os.path.exists(pdf):
                finfo.invoice_analyzed = True
                return False
            res = inv_detect.detect_invoice(pdf, finfo.filename)
            finfo.invoice_detected = bool(res["is_invoice"])
            finfo.invoice_confidence = float(res.get("confidence", 0.0))
            if res["is_invoice"]:
                finfo.document_type = "e_invoice"
            elif res.get("needs_ocr"):
                finfo.document_type = "unknown_scan"
            else:
                finfo.document_type = "document"
            finfo.invoice_analyzed = True
            return finfo.invoice_detected
        except Exception:
            try:
                finfo.invoice_analyzed = True
            except Exception:
                pass
            return False

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

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Chọn tệp tin cần in",
            filetypes=FILE_DIALOG_TYPES,
        )
        if paths:
            self._add_to_queue(list(paths))

    def add_folder(self):
        folder = filedialog.askdirectory(title="Chọn thư mục")
        if not folder:
            return

        if self._include_subfolders.get():
            found = self.pdf_manager.scan_folder(folder, recursive=True)
            if not found:
                messagebox.showinfo("Thông báo", "Không tìm thấy file hỗ trợ trong thư mục.")
                return
            self._add_to_queue(found)
        else:
            paths = filedialog.askopenfilenames(
                title=f"Chọn tệp tin từ: {os.path.basename(folder)}",
                initialdir=folder,
                filetypes=FILE_DIALOG_TYPES,
            )
            if paths:
                self._add_to_queue(list(paths))

    def _add_to_queue(self, paths: list[str]):
        if not paths:
            return
        default_copies = max(1, self._copies_var.get())

        candidates = []
        for p in paths:
            ext = os.path.splitext(p)[1].lower()
            if ext in SUPPORTED_EXTENSIONS and os.path.exists(p):
                candidates.append(p)
        seen = {os.path.abspath(f.original_path) for f in self.file_list}
        valid_paths = dedupe_paths(candidates, seen)

        if not valid_paths:
            self._log("⚠️ Không tìm thấy tệp tin được hỗ trợ.")
            return

        total_to_add = len(valid_paths)
        self.status_lbl.configure(text=f"⚡ Đang nạp {total_to_add} tệp tin...")

        import concurrent.futures
        import threading

        # Snapshot Tk vars ở main thread (thread nền không được chạm Tk)
        _inv_mode = self._invoice_mode_var.get()

        def _bg_loader():
            added_infos = []
            errors = []

            # Parallel metadata parsing
            max_w = min(32, max(4, (os.cpu_count() or 4) * 4))
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_w) as executor:
                futures = [executor.submit(self.pdf_manager.add_file, p, True) for p in valid_paths]
                for p, f in zip(valid_paths, futures):
                    try:
                        info = f.result()
                        info.copies = default_copies
                        added_infos.append(info)
                    except Exception as exc:
                        errors.append(f"{os.path.basename(p)}: {exc}")

            # Nhận diện hóa đơn từng file PDF (tuần tự, nhẹ — text đã có sẵn).
            # _inv_mode đã snapshot ở main thread (Tk var cấm chạm từ nền).
            detected_names = []
            if _inv_mode != INVOICE_OFF:
                for info in added_infos:
                    try:
                        if self._analyze_invoice(info) and info.invoice_detected:
                            detected_names.append(info.filename)
                    except Exception:
                        pass

            def _on_done():
                if not getattr(self, "_is_alive", False):
                    return
                for info in added_infos:
                    self.file_list.append(info)
                    tag = " 🧾 Hóa đơn" if info.invoice_detected else ""
                    self._log(f"+ {info.filename}  ({info.file_type}, {info.page_count} trang){tag}")
                for err in errors:
                    self._log(f"⚠ Lỗi: {err}")

                self._refresh_tree(preserve_selection=True)
                self._update_stats_pill()
                self.status_lbl.configure(text=f"✓ Đã nạp xong {len(added_infos)} tệp tin")

                if added_infos:
                    self._log(f"✓ Đã thêm {len(added_infos)} tệp vào hàng đợi (Tổng: {len(self.file_list)} tệp)")
                    if len(self.file_list) == len(added_infos):
                        children = self.tree.get_children()
                        if children:
                            self.tree.selection_set(children[0])
                            self._on_file_select(None)
                if errors:
                    messagebox.showwarning(
                        "Cảnh báo",
                        f"Không thể thêm {len(errors)} tệp tin:\n\n" + "\n".join(errors[:10]),
                    )

            self._async_queue.put((_on_done, ()))

        threading.Thread(target=_bg_loader, daemon=True).start()

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        indices = sorted(
            (int(self.tree.item(iid, "values")[0]) - 1 for iid in sel),
            reverse=True,
        )
        for i in indices:
            if 0 <= i < len(self.file_list):
                del self.file_list[i]
        self._refresh_tree()
        self._update_stats_pill()
        self._update_queue_summary()
        if hasattr(self, "selected_file_lbl"):
            self.selected_file_lbl.configure(text="📄 (Chưa chọn tệp)", text_color=THEME_COLORS["text_muted"])
        self._log(f"Đã xóa {len(indices)} mục khỏi hàng đợi")

    def remove_all(self):
        if not self.file_list:
            return
        n = len(self.file_list)
        self.file_list.clear()
        self._refresh_tree()
        self._update_stats_pill()
        self._update_queue_summary()
        if hasattr(self, "selected_file_lbl"):
            self.selected_file_lbl.configure(text="📄 (Chưa chọn tệp)", text_color=THEME_COLORS["text_muted"])
        self._log(f"Đã xóa tất cả ({n} mục)")

    def _refresh_tree(self, preserve_selection: bool = False):
        selected_indices = []
        if preserve_selection:
            for item in self.tree.selection():
                try:
                    selected_indices.append(int(self.tree.item(item, "values")[0]) - 1)
                except Exception:
                    pass

        self.tree.delete(*self.tree.get_children())
        items_to_select = []
        for i, f in enumerate(self.file_list):
            icon = STATUS_ICONS.get(f.status, "")
            item_id = self.tree.insert("", "end", values=(
                i + 1,
                f.filename,
                f.file_type,
                f.page_count,
                format_file_size(f.file_size),
                f.pages_display(),
                f.copies,
                f"{icon} {f.status}",
                f.doctype_display(),
                f.duplex_display(),
            ))
            if preserve_selection and i in selected_indices:
                items_to_select.append(item_id)

        if items_to_select:
            self.tree.selection_set(items_to_select)

        self._update_queue_summary()

    def _update_stats_pill(self):
        total_files, _copies, total_pages = queue_summary(self.file_list)
        page_text = f"{total_pages} trang in" if total_pages is not None else "Đang tính..."
        self.stats_lbl.configure(text=f"📦 {total_files} tệp tin  •  📄 {page_text}")
        self._update_queue_summary()

    def _update_queue_summary(self):
        total_files, total_copies, total_pages = queue_summary(self.file_list)
        page_text = f"{total_pages} trang" if total_pages is not None else "Đang tính..."

        if hasattr(self, "queue_summary_lbl"):
            self.queue_summary_lbl.configure(
                text=f"📦 Tổng tài liệu: {total_files}  •  📑 Tổng số bản: {total_copies}  •  📄 Tổng số trang in: {page_text}"
            )

    def _on_file_select(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            if hasattr(self, "selected_file_lbl"):
                self.selected_file_lbl.configure(text="📄 (Chưa chọn tệp)", text_color=THEME_COLORS["text_muted"])
            return
        try:
            idx = int(self.tree.item(sel[0], "values")[0]) - 1
            if 0 <= idx < len(self.file_list):
                finfo = self.file_list[idx]
                if hasattr(self, "selected_file_lbl"):
                    name = finfo.filename if len(finfo.filename) <= 32 else finfo.filename[:31] + "…"
                    self.selected_file_lbl.configure(text=f"📄 {name}", text_color=THEME_COLORS["text"])
                self._copies_var.set(finfo.copies)
                self._sync_page_ui_from_selection()
        except Exception:
            pass

    def move_selected_up(self):
        """Move selected items up one position in the queue."""
        if self._is_printing:
            messagebox.showwarning("Cảnh báo", "Không thể thay đổi thứ tự hàng đợi khi đang in.")
            return

        sel = self.tree.selection()
        if not sel:
            return

        indices = []
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    indices.append(idx)
            except Exception:
                pass

        if not indices:
            return

        new_indices = apply_move(self.file_list, indices, -1)

        self.tree.delete(*self.tree.get_children())
        items_to_select = []
        for i, f in enumerate(self.file_list):
            icon = STATUS_ICONS.get(f.status, "")
            item_id = self.tree.insert("", "end", values=(
                i + 1,
                f.filename,
                f.file_type,
                f.page_count,
                format_file_size(f.file_size),
                f.pages_display(),
                f.copies,
                f"{icon} {f.status}",
                f.doctype_display(),
                f.duplex_display(),
            ))
            if i in new_indices:
                items_to_select.append(item_id)

        if items_to_select:
            self.tree.selection_set(items_to_select)
            self.tree.see(items_to_select[0])

    def move_selected_down(self):
        """Move selected items down one position in the queue."""
        if self._is_printing:
            messagebox.showwarning("Cảnh báo", "Không thể thay đổi thứ tự hàng đợi khi đang in.")
            return

        sel = self.tree.selection()
        if not sel:
            return

        indices = []
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    indices.append(idx)
            except Exception:
                pass

        if not indices:
            return

        new_indices = apply_move(self.file_list, indices, +1)

        self.tree.delete(*self.tree.get_children())
        items_to_select = []
        for i, f in enumerate(self.file_list):
            icon = STATUS_ICONS.get(f.status, "")
            item_id = self.tree.insert("", "end", values=(
                i + 1,
                f.filename,
                f.file_type,
                f.page_count,
                format_file_size(f.file_size),
                f.pages_display(),
                f.copies,
                f"{icon} {f.status}",
                f.doctype_display(),
                f.duplex_display(),
            ))
            if i in new_indices:
                items_to_select.append(item_id)

        if items_to_select:
            self.tree.selection_set(items_to_select)
            self.tree.see(items_to_select[-1])

    def _sort_by_column(self, col_id: str):
        """Sort file queue by clicked column heading with ascending/descending toggle."""
        if self._is_printing:
            messagebox.showwarning("Cảnh báo", "Không thể sắp xếp hàng đợi khi đang in.")
            return

        if not self.file_list:
            return

        is_reverse = self._sort_reverse.get(col_id, False)
        new_reverse = not is_reverse
        self._sort_reverse[col_id] = new_reverse

        key_funcs = {
            "stt": lambda f: 0,
            "filename": lambda f: f.filename.lower(),
            "filetype": lambda f: f.file_type.lower(),
            "pages": lambda f: int(f.page_count) if isinstance(f.page_count, int) else 0,
            "size": lambda f: f.file_size,
            "pagesel": lambda f: f.pages_display(),
            "copies": lambda f: f.copies,
            "status": lambda f: str(f.status),
            "doctype": lambda f: f.document_type,
            "duplex": lambda f: f.duplex_display(),
        }

        k_fn = key_funcs.get(col_id, lambda f: f.filename.lower())
        self.file_list.sort(key=k_fn, reverse=new_reverse)

        headings_map = {
            "stt": "#",
            "filename": "Tên Tệp Tin",
            "filetype": "Định Dạng",
            "pages": "Số Trang",
            "size": "Kích Thước",
            "pagesel": "Trang In",
            "copies": "Số Bản (✎)",
            "status": "Trạng Thái",
            "doctype": "Loại",
            "duplex": "2 Mặt",
        }
        for cid, label in headings_map.items():
            if cid == col_id:
                indicator = " ▼" if new_reverse else " ▲"
                self.tree.heading(cid, text=f"{label}{indicator}")
            else:
                self.tree.heading(cid, text=f"{label} ⇕")

        self._refresh_tree(preserve_selection=True)
        dir_text = "Giảm dần (Z-A)" if new_reverse else "Tăng dần (A-Z)"
        self._log(f"Đã sắp xếp danh sách theo [{headings_map.get(col_id, col_id)}] — {dir_text}")

    def save_queue_session(self):
        """Save current queue file list and copy counts to a JSON session file."""
        if not self.file_list:
            messagebox.showinfo("Lưu danh sách", "Hàng đợi đang trống, không có tệp tin nào để lưu.")
            return

        path = filedialog.asksaveasfilename(
            title="Lưu danh sách hàng đợi in",
            defaultextension=".json",
            filetypes=[("Tệp phiên in (*.json)", "*.json"), ("Tất cả tệp (*.*)", "*.*")],
            initialfile=f"danh_sach_in_{datetime.now():%Y%m%d_%H%M%S}.json",
        )
        if not path:
            return

        data = []
        for f in self.file_list:
            data.append({
                "source_path": f.original_path if hasattr(f, "original_path") else f.pdf_path,
                "copies": f.copies,
                "filename": f.filename,
                "file_type": f.file_type,
                "pages": {"mode": getattr(f, "page_mode", "all"),
                          "value": getattr(f, "page_range_text", "")},
                "document_type": getattr(f, "document_type", "document"),
                "invoice_detected": bool(getattr(f, "invoice_detected", False)),
                "invoice_confidence": float(getattr(f, "invoice_confidence", 0.0) or 0.0),
                "duplex_mode": getattr(f, "duplex_mode", None),
                "duplex_auto": bool(getattr(f, "duplex_auto", False)),
                "duplex_override": bool(getattr(f, "duplex_override", False)),
            })

        try:
            with open(path, "w", encoding="utf-8") as fp:
                json.dump({"version": "2.0", "files": data}, fp, indent=2, ensure_ascii=False)
            self._log(f"✓ Đã lưu danh sách {len(data)} tệp tin vào: {path}")
            messagebox.showinfo("Thành công", f"Đã lưu danh sách {len(data)} tệp tin thành công!")
        except Exception as exc:
            messagebox.showerror("Lỗi", f"Không thể lưu danh sách:\n{exc}")

    def load_queue_session(self):
        """Load queue files and copy counts from a saved JSON session file."""
        if self._is_printing:
            messagebox.showwarning("Cảnh báo", "Không thể nạp danh sách khi đang in.")
            return

        path = filedialog.askopenfilename(
            title="Mở danh sách hàng đợi in",
            filetypes=[("Tệp phiên in (*.json)", "*.json"), ("Tất cả tệp (*.*)", "*.*")],
        )
        if not path:
            return

        try:
            with open(path, "r", encoding="utf-8") as fp:
                session_data = json.load(fp)

            items = session_data.get("files", [])
            if not items:
                messagebox.showinfo("Thông báo", "Tệp danh sách không chứa tệp tin nào.")
                return

            self.status_lbl.configure(text=f"⚡ Đang nạp danh sách {len(items)} tệp...")

            import concurrent.futures
            import threading

            def _bg_session_loader():
                added_infos = []
                missing = []

                # Filter existing files
                valid_items = []
                for item in items:
                    src_path = item.get("source_path")
                    if not src_path or not os.path.exists(src_path):
                        missing.append(item.get("filename", src_path or "Unknown"))
                    else:
                        valid_items.append(item)

                max_w = min(32, max(4, (os.cpu_count() or 4) * 4))
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_w) as executor:
                    futures = [executor.submit(self.pdf_manager.add_file, it.get("source_path"), True) for it in valid_items]
                    for it, f in zip(valid_items, futures):
                        src_path = it.get("source_path")
                        copies = it.get("copies", 1)
                        pages = it.get("pages", {}) or {}
                        try:
                            info = f.result()
                            info.copies = max(1, min(999, int(copies)))
                            mode = str(pages.get("mode", "all") or "all").lower()
                            if mode not in ("all", "custom", "odd", "even"):
                                mode = "all"
                            info.page_mode = mode
                            info.page_range_text = str(pages.get("value", "") or "")
                            info.document_type = str(it.get("document_type", "document") or "document")
                            info.invoice_detected = bool(it.get("invoice_detected", False))
                            try:
                                info.invoice_confidence = float(it.get("invoice_confidence", 0.0) or 0.0)
                            except Exception:
                                info.invoice_confidence = 0.0
                            dm = it.get("duplex_mode")
                            info.duplex_mode = dm if dm in ("simplex", "long", "short") else None
                            info.duplex_auto = bool(it.get("duplex_auto", False))
                            info.duplex_override = bool(it.get("duplex_override", False))
                            info.invoice_analyzed = True
                            added_infos.append(info)
                        except Exception as exc:
                            missing.append(f"{os.path.basename(src_path)} ({exc})")

                def _on_done():
                    if not getattr(self, "_is_alive", False):
                        return
                    for info in added_infos:
                        self.file_list.append(info)
                    self._refresh_tree(preserve_selection=True)
                    self._update_stats_pill()
                    self.status_lbl.configure(text=f"✓ Đã nạp xong {len(added_infos)} tệp từ danh sách")

                    self._log(f"Đã mở danh sách: nạp thành công {len(added_infos)} tệp từ {os.path.basename(path)}")
                    if missing:
                        self._log(f"⚠️ {len(missing)} tệp trong danh sách không tìm thấy hoặc bị lỗi")
                        messagebox.showwarning(
                            "Cảnh báo",
                            f"Đã nạp {len(added_infos)} tệp tin.\n{len(missing)} tệp không tìm thấy trên máy:\n\n" + "\n".join(missing[:8]),
                        )

                self._async_queue.put((_on_done, ()))

            threading.Thread(target=_bg_session_loader, daemon=True).start()
        except Exception as exc:
            messagebox.showerror("Lỗi", f"Không thể đọc tệp danh sách:\n{exc}")

    # ── Rapid Per-File Copy Editing System ───────────────────────────

    def _on_tree_click(self, event):
        """Single click on copies column (#7) edits copies; on Trang In (#6) edits pages."""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            col = self.tree.identify_column(event.x)
            if col == "#7":
                item = self.tree.identify_row(event.y)
                if item:
                    self.after(40, lambda: self._start_inline_edit(item, "#7"))
            elif col == "#6":
                item = self.tree.identify_row(event.y)
                if item:
                    self.after(40, lambda: self._start_pages_inline_edit(item))

    def _on_tree_double_click(self, event):
        """Double click on copies (#7) edits copies, on Trang In (#6) edits pages, elsewhere opens preview."""
        region = self.tree.identify("region", event.x, event.y)
        if region in ("cell", "tree"):
            item = self.tree.identify_row(event.y)
            if item:
                col = self.tree.identify_column(event.x)
                if col == "#7":
                    self._start_inline_edit(item, "#7")
                elif col == "#6":
                    self._start_pages_inline_edit(item)
                else:
                    self._preview_selected()

    def _start_inline_edit(self, item, col_id="#7"):
        """Spawn an in-place entry widget on the treeview cell for editing copies."""
        if not item:
            return
        bbox = self.tree.bbox(item, col_id)
        if not bbox:
            return

        x, y, w, h = bbox
        try:
            values = self.tree.item(item, "values")
            idx = int(values[0]) - 1
        except Exception:
            return

        if idx < 0 or idx >= len(self.file_list):
            return

        f = self.file_list[idx]
        # Check if file is currently printing or already printed
        if f.status == FileStatus.PRINTING:
            messagebox.showwarning(
                "Không thể chỉnh sửa",
                f"Tệp tin '{f.filename}' đang trong quá trình in ấn.\nBạn không thể sửa số bản in lúc này.",
            )
            return
        if f.status == FileStatus.PRINTED:
            messagebox.showwarning(
                "Không thể chỉnh sửa",
                f"Tệp tin '{f.filename}' đã hoàn thành in.\nKhông thể sửa đổi số bản in của tệp đã in.",
            )
            return

        if self._edit_entry is not None:
            try:
                self._edit_entry.destroy()
            except Exception:
                pass
            self._edit_entry = None

        entry = tk.Entry(
            self.tree,
            justify="center",
            font=("Segoe UI", 10, "bold"),
            bd=2,
            relief="solid",
            highlightthickness=1,
            highlightcolor="#2563EB",
            bg="#EFF6FF",
            fg="#1E40AF",
        )
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, str(f.copies))
        entry.select_range(0, "end")
        entry.focus_set()
        self._edit_entry = entry

        def _commit_and_move(direction=0):
            try:
                val = int(entry.get().strip())
                val = max(1, min(999, val))
            except (ValueError, AttributeError):
                val = self.file_list[idx].copies

            self.file_list[idx].copies = val
            if self._edit_entry == entry:
                entry.destroy()
                self._edit_entry = None

            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._copies_var.set(val)

            # Fast navigation to next/previous available (unprinted) row
            if direction != 0:
                children = self.tree.get_children()
                try:
                    curr_pos = children.index(item)
                    next_pos = curr_pos + direction
                    while 0 <= next_pos < len(children):
                        cand_item = children[next_pos]
                        cand_idx = int(self.tree.item(cand_item, "values")[0]) - 1
                        if 0 <= cand_idx < len(self.file_list):
                            cand_f = self.file_list[cand_idx]
                            if cand_f.status not in (FileStatus.PRINTING, FileStatus.PRINTED):
                                self.tree.selection_set(cand_item)
                                self.tree.see(cand_item)
                                self.after(30, lambda c=cand_item: self._start_inline_edit(c, col_id))
                                break
                        next_pos += direction
                except Exception:
                    pass

        def _cancel(_e=None):
            if self._edit_entry == entry:
                entry.destroy()
                self._edit_entry = None

        entry.bind("<Return>", lambda e: _commit_and_move(1))
        entry.bind("<Tab>", lambda e: (_commit_and_move(1), "break")[1])
        entry.bind("<Shift-Tab>", lambda e: (_commit_and_move(-1), "break")[1])
        entry.bind("<Down>", lambda e: _commit_and_move(1))
        entry.bind("<Up>", lambda e: _commit_and_move(-1))
        entry.bind("<Escape>", _cancel)
        entry.bind("<FocusOut>", lambda e: _commit_and_move(0))

    def _start_pages_inline_edit(self, item):
        """Popup editor for the 'Trang In' cell: mode dropdown + custom range entry."""
        if not item:
            return
        try:
            values = self.tree.item(item, "values")
            idx = int(values[0]) - 1
        except Exception:
            return
        if idx < 0 or idx >= len(self.file_list):
            return
        f = self.file_list[idx]
        if is_edit_locked(f.status):
            messagebox.showwarning(
                "Không thể chỉnh sửa",
                f"Tệp tin '{f.filename}' đang trong quá trình in hoặc đã in xong.\nKhông thể đổi trang in lúc này.",
            )
            return

        # Close any stale copies editor
        if self._edit_entry is not None:
            try:
                self._edit_entry.destroy()
            except Exception:
                pass
            self._edit_entry = None

        # Position popup near the clicked cell
        try:
            bbox = self.tree.bbox(item, "#6")
            if bbox:
                x, y, _w, h = bbox
                px = self.tree.winfo_rootx() + x
                py = self.tree.winfo_rooty() + y + h
            else:
                px, py = self.tree.winfo_rootx() + 100, self.tree.winfo_rooty() + 60
        except Exception:
            px, py = self.tree.winfo_rootx() + 100, self.tree.winfo_rooty() + 60

        dlg = ctk.CTkToplevel(self)
        dlg.title(f"Trang in — {f.filename}")
        dlg.geometry(f"320x210+{px}+{py}")
        dlg.transient(self)
        dlg.resizable(False, False)
        try:
            dlg.grab_set()
        except Exception:
            pass

        ctk.CTkLabel(
            dlg, text=f"📄 {f.filename}"[:42],
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack(padx=14, pady=(12, 2), anchor="w")
        ctk.CTkLabel(
            dlg, text=f"Tài liệu có {f.page_count} trang. Chọn chế độ in:",
            font=ctk.CTkFont(size=11),
            text_color=THEME_COLORS["text_muted"],
        ).pack(padx=14, pady=(0, 8), anchor="w")

        mode_var = tk.StringVar(value=PAGE_LABEL_BY_MODE.get(f.page_mode, PAGE_RANGE_ALL))
        val_var = tk.StringVar(value=f.page_range_text)

        combo = ctk.CTkComboBox(
            dlg, values=PAGE_RANGE_OPTIONS, variable=mode_var,
            width=292, height=30, state="readonly",
            font=ctk.CTkFont(size=11),
        )
        combo.pack(padx=14, pady=(0, 6))

        entry = ctk.CTkEntry(
            dlg, textvariable=val_var, width=292, height=30,
            placeholder_text="Ví dụ: 1,3,5-8,12",
            font=ctk.CTkFont(size=11),
        )

        def _toggle_entry(_v=None):
            if mode_var.get() == PAGE_RANGE_CUSTOM:
                entry.pack(padx=14, pady=(0, 6))
                entry.focus_set()
            else:
                try:
                    entry.pack_forget()
                except Exception:
                    pass

        combo.configure(command=_toggle_entry)
        _toggle_entry()

        btn_row = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_row.pack(padx=14, pady=(4, 12), fill="x")

        def _on_ok():
            mode = PAGE_MODE_BY_LABEL.get(mode_var.get(), "all")
            raw = val_var.get().strip()
            if mode == "custom":
                try:
                    resolve_page_selection("custom", raw, f.page_count)
                except ValueError as exc:
                    messagebox.showwarning("Trang in không hợp lệ",
                                           f"{f.filename} ({f.page_count} trang):\n{exc}",
                                           parent=dlg)
                    return
            f.page_mode = mode
            f.page_range_text = raw if mode == "custom" else ""
            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._sync_page_ui_from_selection()
            self._log(f"Đã đặt trang in [{f.pages_display()}] cho '{f.filename}'")
            self._save_config()
            try:
                dlg.destroy()
            except Exception:
                pass

        ctk.CTkButton(
            btn_row, text="✔ Đồng ý", command=_on_ok, height=30,
            fg_color=THEME_COLORS["primary"],
            hover_color=THEME_COLORS["primary_hover"],
            text_color="#FFFFFF", text_color_disabled="#FFFFFF",
            font=ctk.CTkFont(size=11, weight="bold"),
        ).pack(side="left", expand=True, fill="x", padx=(0, 4))
        ctk.CTkButton(
            btn_row, text="Hủy", command=dlg.destroy, height=30,
            fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            text_color_disabled=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(size=11),
        ).pack(side="left", expand=True, fill="x", padx=(4, 0))

        dlg.bind("<Escape>", lambda e: dlg.destroy())
        entry.bind("<Return>", lambda e: _on_ok())

    def _edit_selected_copies(self):
        """Dedicated edit button handler — verifies printing state and triggers editor."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Chỉnh sửa số bản in", "Vui lòng chọn ít nhất một tệp tin trong danh sách để sửa số bản in.")
            return

        if len(sel) == 1:
            item = sel[0]
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
            except Exception:
                return
            if 0 <= idx < len(self.file_list):
                f = self.file_list[idx]
                if f.status == FileStatus.PRINTING:
                    messagebox.showwarning(
                        "Không thể chỉnh sửa",
                        f"Tệp tin '{f.filename}' đang trong quá trình in ấn.\nBạn không thể sửa số bản in lúc này.",
                    )
                    return
                if f.status == FileStatus.PRINTED:
                    messagebox.showwarning(
                        "Không thể chỉnh sửa",
                        f"Tệp tin '{f.filename}' đã in xong.\nKhông thể sửa đổi số bản in của tệp đã in.",
                    )
                    return
                self._start_inline_edit(item, "#7")
        else:
            self._prompt_change_copies_dialog()

    def _on_tree_keypress(self, event):
        """Handle quick keyboard shortcuts for adjusting copies on selected items."""
        if self._edit_entry is not None:
            return

        sel = self.tree.selection()
        if not sel:
            return

        # Plus / Equal key -> Increase copies by 1
        if event.char in ("+", "="):
            self._change_selected_copies(delta=1)
            return "break"
        # Minus / Underscore key -> Decrease copies by 1
        elif event.char in ("-", "_"):
            self._change_selected_copies(delta=-1)
            return "break"
        # Number keys 1-9 -> Set exact copy count
        elif event.char in "123456789":
            copies = int(event.char)
            self._set_selected_copies(copies)
            return "break"
        # F2, Return, Space -> Open inline editor for selected item
        elif event.keysym in ("F2", "Return", "space"):
            self._edit_selected_copies()
            return "break"

    def _on_tree_right_click(self, event):
        """Right-click popup context menu for fast copy modifications."""
        item = self.tree.identify_row(event.y)
        if item and item not in self.tree.selection():
            self.tree.selection_set(item)

        sel = self.tree.selection()
        if not sel:
            return

        menu = tk.Menu(self, tearoff=0, font=("Segoe UI", 10))
        n = len(sel)
        suffix = f" ({n} tệp)" if n > 1 else ""

        menu.add_command(label="🔍 Xem trước chi tiết (Mở cửa sổ lớn)...", command=self._preview_selected)
        menu.add_separator()
        menu.add_command(label=f"✏️ Sửa số bản in...{suffix}", command=self._edit_selected_copies)
        if n == 1:
            menu.add_command(label="📄 Đặt trang in riêng...",
                             command=lambda: self._start_pages_inline_edit(sel[0]))
        menu.add_separator()
        menu.add_command(label="🎯 In các mục đang chọn", command=self.start_print_selected)
        menu.add_separator()
        menu.add_command(label="➕ Tăng +1 bản (+)", command=lambda: self._change_selected_copies(1))
        menu.add_command(label="➖ Giảm -1 bản (-)", command=lambda: self._change_selected_copies(-1))

        # Preset submenu
        preset_menu = tk.Menu(menu, tearoff=0, font=("Segoe UI", 10))
        for count in [1, 2, 3, 4, 5, 10, 20]:
            preset_menu.add_command(
                label=f"{count} bản",
                command=lambda c=count: self._set_selected_copies(c),
            )
        menu.add_cascade(label="📋 Đặt nhanh số bản", menu=preset_menu)

        menu.add_separator()
        menu.add_command(label="▲ Di chuyển lên (Ctrl+Up)", command=self.move_selected_up)
        menu.add_command(label="▼ Di chuyển xuống (Ctrl+Down)", command=self.move_selected_down)
        menu.add_separator()
        menu.add_command(label="💾 Lưu danh sách (JSON)...", command=self.save_queue_session)
        menu.add_command(label="📂 Mở danh sách (JSON)...", command=self.load_queue_session)
        menu.add_separator()
        menu.add_command(label="🗑️ Xóa khỏi hàng đợi (Del)", command=self.remove_selected)

        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _change_selected_copies(self, delta: int):
        sel = self.tree.selection()
        if not sel:
            return
        last_val = 1
        skipped_names = []
        updated_count = 0
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    f = self.file_list[idx]
                    if f.status in (FileStatus.PRINTING, FileStatus.PRINTED):
                        skipped_names.append(f.filename)
                        continue
                    new_val = max(1, min(999, f.copies + delta))
                    f.copies = new_val
                    last_val = new_val
                    updated_count += 1
            except Exception:
                pass

        if skipped_names:
            self._log(f"⚠️ Bỏ qua {len(skipped_names)} tệp đang in / đã in xong khi chỉnh số bản")
            if updated_count == 0:
                messagebox.showwarning(
                    "Không thể chỉnh sửa",
                    "Các tệp tin được chọn đang trong quá trình in hoặc đã hoàn thành in,\nkhông thể sửa đổi số bản in.",
                )

        if updated_count > 0:
            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._copies_var.set(last_val)

    def _set_selected_copies(self, copies: int):
        sel = self.tree.selection()
        if not sel:
            return
        copies = max(1, min(999, copies))
        skipped_names = []
        updated_count = 0
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    f = self.file_list[idx]
                    if f.status in (FileStatus.PRINTING, FileStatus.PRINTED):
                        skipped_names.append(f.filename)
                        continue
                    f.copies = copies
                    updated_count += 1
            except Exception:
                pass

        if skipped_names:
            self._log(f"⚠️ Bỏ qua {len(skipped_names)} tệp đang in / đã in xong khi đặt số bản")
            if updated_count == 0:
                messagebox.showwarning(
                    "Không thể chỉnh sửa",
                    "Các tệp tin được chọn đang trong quá trình in hoặc đã hoàn thành in,\nkhông thể sửa đổi số bản in.",
                )

        if updated_count > 0:
            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._copies_var.set(copies)
            self._log(f"Đã đặt {copies} bản in cho {updated_count} tệp tin")

    def _prompt_change_copies_dialog(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Thông báo", "Vui lòng chọn ít nhất một tệp tin trong danh sách để đổi số bản in.")
            return

        valid_items = []
        skipped_names = []
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    f = self.file_list[idx]
                    if f.status in (FileStatus.PRINTING, FileStatus.PRINTED):
                        skipped_names.append(f.filename)
                    else:
                        valid_items.append(f)
            except Exception:
                pass

        if not valid_items:
            messagebox.showwarning(
                "Không thể chỉnh sửa",
                "Tất cả các tệp tin được chọn đều đang trong quá trình in hoặc đã in xong,\nkhông thể sửa đổi số bản in.",
            )
            return

        current_val = valid_items[0].copies
        count_text = f"{len(valid_items)} tệp tin có thể chỉnh sửa" if len(valid_items) > 1 else f"tệp '{valid_items[0].filename}'"
        dialog = ctk.CTkInputDialog(
            text=f"Nhập số bản in mới cho {count_text} (1 - 999):",
            title="Chỉnh Sửa Số Bản In",
        )
        val_str = dialog.get_input()
        if val_str is not None and val_str.strip():
            try:
                val = int(val_str.strip())
                if 1 <= val <= 999:
                    self._set_selected_copies(val)
                else:
                    messagebox.showwarning("Cảnh báo", "Số bản in phải từ 1 đến 999.")
            except ValueError:
                messagebox.showwarning("Cảnh báo", "Vui lòng nhập số nguyên hợp lệ.")

    def _preview_selected(self):
        """Open the large dedicated preview modal for the selected file."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Xem trước", "Vui lòng chọn một tệp tin trong hàng đợi để xem trước.")
            return
        try:
            idx = int(self.tree.item(sel[0], "values")[0]) - 1
            if 0 <= idx < len(self.file_list):
                finfo = self.file_list[idx]
                # Tái dùng 1 cửa sổ preview để tránh rò rỉ hàng chục Toplevel
                win = getattr(self, "_preview_win", None)
                try:
                    if win is not None and win.winfo_exists():
                        win.load_target(finfo, 0)
                        win.lift()
                        win.focus_force()
                        return
                except Exception:
                    pass
                self._preview_win = PreviewWindow(self, self.pdf_manager, finfo)
        except Exception as exc:
            print(f"Error previewing file: {exc}")

    # ═══════════════════════════════════════════════════════════════════
    #  PRINTER & OPTIONS
    # ═══════════════════════════════════════════════════════════════════

    def load_printers(self):
        """Asynchronously load printers from Windows in a background thread."""
        self.printer_combo.configure(values=["(Đang tải danh sách máy in...)"])
        self._printer_var.set("(Đang tải danh sách máy in...)")
        self.printer_status_lbl.configure(text="⏳ Đang quét máy in...")

        def _worker():
            try:
                printers = PrinterManager.get_printers()
                default = PrinterManager.get_default_printer()
                self._async_queue.put((self._on_printers_loaded, (printers, default)))
            except Exception as exc:
                print(f"Error loading printers: {exc}")

        import threading
        threading.Thread(target=_worker, daemon=True).start()

    def _on_printers_loaded(self, printers: list[str], default: str):
        self._available_printers = printers
        if not printers:
            self.printer_combo.configure(values=["(Không tìm thấy máy in)"])
            self._printer_var.set("(Không tìm thấy máy in)")
            self.printer_status_lbl.configure(text="🔴 Không có máy in")
            self.failover_combo.configure(values=["(Không dùng)"])
            self._failover_printer_var.set("(Không dùng)")
            return

        self.printer_combo.configure(values=printers)
        failover_list = ["(Không dùng)"] + printers
        self.failover_combo.configure(values=failover_list)

        if getattr(self, "_saved_failover", "") in failover_list:
            self._failover_printer_var.set(self._saved_failover)
        else:
            self._failover_printer_var.set("(Không dùng)")

        if getattr(self, "_saved_printer", "") and self._saved_printer in printers:
            selected = self._saved_printer
        else:
            selected = default if default in printers else printers[0]
        self._printer_var.set(selected)
        self._update_printer_status()

    def _open_parallel_printers_dialog(self):
        """Opens a modal dialog to select multiple printers for parallel load-balanced printing."""
        if not self._available_printers:
            messagebox.showinfo("In song song", "Chưa có danh sách máy in sẵn sàng.")
            return

        dialog = ctk.CTkToplevel(self)
        dialog.title("🖨️ In Song Song Ra Nhiều Máy In")
        dialog.geometry("460x400")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)

        ctk.CTkLabel(
            dialog, text="🖨️ Phân Phối Tải Song Song (Multi-Printer)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=THEME_COLORS["text"],
        ).pack(padx=16, pady=(14, 4), anchor="w")

        ctk.CTkLabel(
            dialog,
            text="Chọn các máy in để chia đều danh sách tệp tin và in đồng thời:\n(Nếu chọn 1 máy hoặc hủy chọn, hệ thống sẽ in tuần tự bình thường)",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=THEME_COLORS["text_muted"],
            justify="left",
        ).pack(padx=16, pady=(0, 10), anchor="w")

        scroll_frame = ctk.CTkScrollableFrame(
            dialog, height=190, fg_color=THEME_COLORS["card_alt"], corner_radius=8
        )
        scroll_frame.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        cb_vars = {}
        for p in self._available_printers:
            is_checked = (p in self._parallel_printers) or (not self._parallel_printers and p == self._printer_var.get())
            var = tk.BooleanVar(value=is_checked)
            cb_vars[p] = var
            ctk.CTkCheckBox(
                scroll_frame, text=p, variable=var,
                font=ctk.CTkFont(family="Segoe UI", size=11),
                text_color=THEME_COLORS["text"],
                checkbox_width=16, checkbox_height=16,
                corner_radius=4,
            ).pack(anchor="w", padx=10, pady=4)

        def _confirm():
            chosen = [p for p, v in cb_vars.items() if v.get()]
            if len(chosen) >= 2:
                self._parallel_printers = chosen
                self.parallel_status_lbl.configure(
                    text=f"✓ Đang chọn {len(chosen)} máy in",
                    text_color=THEME_COLORS["success"][0],
                )
                self._log(f"Đã kích hoạt in song song trên {len(chosen)} máy in: {', '.join(chosen)}")
            elif len(chosen) == 1:
                self._parallel_printers = []
                self._printer_var.set(chosen[0])
                self.parallel_status_lbl.configure(text="")
                self._log(f"Đã chọn máy in: {chosen[0]}")
            else:
                self._parallel_printers = []
                self.parallel_status_lbl.configure(text="")
            self._save_config()
            dialog.destroy()

        def _clear():
            self._parallel_printers = []
            self.parallel_status_lbl.configure(text="")
            self._save_config()
            dialog.destroy()

        btn_box = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_box.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkButton(
            btn_box, text="Xác Nhận", width=100, height=32,
            command=_confirm, fg_color=THEME_COLORS["primary"],
            hover_color=THEME_COLORS["primary_hover"],
            font=ctk.CTkFont(family="Segoe UI", size=11, weight="bold"),
            corner_radius=6,
        ).pack(side="right", padx=(6, 0))

        ctk.CTkButton(
            btn_box, text="Hủy / Mặc định", width=110, height=32,
            command=_clear, fg_color=THEME_COLORS["btn_secondary"],
            hover_color=THEME_COLORS["btn_secondary_hover"],
            text_color=THEME_COLORS["btn_secondary_text"],
            font=ctk.CTkFont(family="Segoe UI", size=11),
            corner_radius=6,
        ).pack(side="right")

    def refresh_printers(self):
        self.load_printers()
        self._log("Đang làm mới danh sách máy in...")

    def _on_printer_change(self, _v=None):
        self._update_printer_status()
        self._save_config()

    def _update_printer_status(self):
        name = self._printer_var.get()
        if not name or name.startswith("("):
            self.printer_status_lbl.configure(text="")
            return

        def _check_worker(_name=name):
            try:
                text, is_ready = PrinterManager.get_printer_status(_name)
                has_duplex = PrinterManager.supports_duplex(_name)
                self._async_queue.put((self._on_status_checked, (_name, text, is_ready, has_duplex)))
            except Exception as exc:
                print(f"Error checking printer status: {exc}")

        import threading
        threading.Thread(target=_check_worker, daemon=True).start()

    def _on_status_checked(self, name: str, text: str, is_ready: bool, has_duplex: bool):
        if self._printer_var.get() != name:
            return
        status_icon = "🟢" if is_ready else "🔴"
        self.printer_status_lbl.configure(text=f"{status_icon} {text}")
        if not has_duplex:
            self.duplex_warn_lbl.configure(text="⚠️ Không hỗ trợ in 2 mặt")
        else:
            self.duplex_warn_lbl.configure(text="")

    def _inc_copies(self):
        try:
            v = int(self._copies_var.get())
        except Exception:
            v = 1
        if v < 999:
            new_v = v + 1
            self._copies_var.set(new_v)
            if self.tree.selection():
                self._set_selected_copies(new_v)
            self._save_config()

    def _dec_copies(self):
        try:
            v = int(self._copies_var.get())
        except Exception:
            v = 1
        if v > 1:
            new_v = v - 1
            self._copies_var.set(new_v)
            if self.tree.selection():
                self._set_selected_copies(new_v)
            self._save_config()

    def _on_copies_entry_enter(self, _event=None):
        try:
            val = int(self._copies_var.get())
            val = max(1, min(999, val))
            self._copies_var.set(val)
            if self.tree.selection():
                self._set_selected_copies(val)
            self._save_config()
        except Exception:
            pass

    def _apply_copies_to_selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Thông báo", "Vui lòng chọn các tệp tin trong hàng đợi trước khi áp dụng.")
            return
        try:
            val = int(self._copies_var.get())
            val = max(1, min(999, val))
            self._copies_var.set(val)
            self._set_selected_copies(val)
            self._save_config()
        except Exception:
            pass

    def _apply_default_copies_to_all(self):
        try:
            c = max(1, min(999, int(self._default_copies_var.get())))
        except Exception:
            c = 1
        self._default_copies_var.set(c)
        updated = 0
        skipped = 0
        for f in self.file_list:
            if f.status in (FileStatus.PRINTING, FileStatus.PRINTED):
                skipped += 1
                continue
            f.copies = c
            updated += 1
        self._refresh_tree(preserve_selection=True)
        self._update_stats_pill()
        self._update_queue_summary()
        msg = f"Đã áp dụng số bản mặc định ({c} bản) cho {updated} tệp tin"
        if skipped > 0:
            msg += f" (Bỏ qua {skipped} tệp đang in/đã in)"
        self._log(msg)
        self._save_config()

    def _selected_indices(self) -> list[int]:
        """Queue indices from current tree selection (in selection order)."""
        indices = []
        for item in self.tree.selection():
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    indices.append(idx)
            except Exception:
                pass
        return indices

    def _editable_indices(self, indices: list[int]) -> tuple[list[int], list[str]]:
        """Split into editable indices + skipped filenames (printing/printed)."""
        ok, skipped = [], []
        for i in indices:
            f = self.file_list[i]
            if is_edit_locked(f.status):
                skipped.append(f.filename)
            else:
                ok.append(i)
        return ok, skipped

    def _sync_page_ui_from_selection(self):
        """Show selected file's page setting in the settings panel."""
        if getattr(self, "_syncing_page_ui", False):
            return
        self._syncing_page_ui = True
        try:
            indices = self._selected_indices()
            if not indices:
                if hasattr(self, "page_file_lbl"):
                    self.page_file_lbl.configure(text="📄 Chưa chọn tệp tin")
                self._sync_doc_ui(None)
                return
            first = self.file_list[indices[0]]
            if len(indices) == 1:
                label = f"📄 {first.filename}"
            else:
                label = f"📄 {len(indices)} tệp đã chọn (hiển thị file đầu: {first.filename})"
            if hasattr(self, "page_file_lbl"):
                self.page_file_lbl.configure(text=label if len(label) <= 52 else label[:51] + "…")
            self._page_range_var.set(PAGE_LABEL_BY_MODE.get(first.page_mode, PAGE_RANGE_ALL))
            self._custom_pages_var.set(first.page_range_text)
            self._refresh_custom_pages_visibility()
            self._sync_doc_ui(first)
        finally:
            self._syncing_page_ui = False

    def _sync_doc_ui(self, finfo=None):
        """Hiển thị nhận diện + nút ghi đè duplex cho file đang chọn (§9-10)."""
        try:
            lbl = getattr(self, "doc_info_lbl", None)
            btn = getattr(self, "btn_duplex_reset", None)
            if lbl is None:
                return
            if finfo is None:
                idxs = self._selected_indices()
                finfo = self.file_list[idxs[0]] if idxs else None
            if finfo is None:
                lbl.configure(text="")
                if btn is not None:
                    btn.pack_forget()
                return
            if finfo.document_type == "e_invoice":
                txt = (f"🧾 Hóa đơn điện tử • {int(finfo.invoice_confidence * 100)}% • "
                       f"{finfo.page_count} trang")
                if finfo.duplex_auto and finfo.duplex_mode == "long":
                    txt += " • ✓ Tự động 2 mặt"
                lbl.configure(text=txt)
            elif finfo.document_type == "unknown_scan":
                lbl.configure(text="🔍 Bản scan – chưa xác định loại")
            else:
                lbl.configure(text="📄 Tài liệu thường")
            if btn is not None:
                if finfo.duplex_auto:
                    btn.pack(side="left", padx=(8, 0))
                else:
                    btn.pack_forget()
        except Exception:
            pass

    def _reset_file_duplex(self):
        """Người dùng ghi đè: file theo mặc định chung, không tự áp lại (§10)."""
        for i in self._selected_indices():
            f = self.file_list[i]
            if is_edit_locked(f.status):
                continue
            f.duplex_mode = None
            f.duplex_auto = False
            f.duplex_override = True
        self._refresh_tree(preserve_selection=True)
        self._sync_doc_ui()
        self._log("Đã chuyển file chọn về duplex mặc định chung (ghi đè nhận diện)")
        self._save_config()

    def _refresh_custom_pages_visibility(self):
        if self._page_range_var.get() == PAGE_RANGE_CUSTOM:
            if not self._custom_pages_visible:
                try:
                    self._pg_custom_row.grid()
                except Exception:
                    pass
                if not getattr(self, "_custom_pages_packed", False):
                    self.custom_pages_entry.pack(fill="x")
                    self._custom_pages_packed = True
                self._custom_pages_visible = True
        else:
            if self._custom_pages_visible:
                try:
                    self._pg_custom_row.grid_remove()
                except Exception:
                    pass
                self._custom_pages_visible = False

    def _on_page_range_change(self, _v=None):
        self._refresh_custom_pages_visibility()
        if getattr(self, "_syncing_page_ui", False):
            return
        mode = PAGE_MODE_BY_LABEL.get(self._page_range_var.get(), "all")
        indices = self._selected_indices()
        if not indices:
            self._save_config()
            return
        editable, skipped = self._editable_indices(indices)
        if mode == "custom":
            # Chờ nhập ô tùy chọn rồi Enter — không áp vội giá trị cũ
            self._save_config()
            return
        for i in editable:
            self.file_list[i].page_mode = mode
            self.file_list[i].page_range_text = ""
        if editable:
            self._refresh_tree(preserve_selection=True)
            self._update_stats_pill()
            self._log(f"Đã đặt trang in [{self._page_range_var.get()}] cho {len(editable)} tệp tin")
        if skipped:
            self._log(f"⚠️ Bỏ qua {len(skipped)} tệp đang in / đã in xong khi đặt trang in")
        self._save_config()

    def _commit_custom_pages(self):
        """Validate + apply custom range text to selected editable files."""
        if getattr(self, "_syncing_page_ui", False):
            return
        if self._page_range_var.get() != PAGE_RANGE_CUSTOM:
            return
        raw = self._custom_pages_var.get().strip()
        indices = self._selected_indices()
        if not indices:
            return
        editable, skipped = self._editable_indices(indices)
        if not editable:
            if skipped:
                messagebox.showwarning("Không thể chỉnh sửa", "Các tệp đang in / đã in xong, không thể đổi trang in.")
            return
        errors = []
        for i in editable:
            f = self.file_list[i]
            try:
                resolve_page_selection("custom", raw, f.page_count)
            except ValueError as exc:
                errors.append(f"{f.filename}: {exc}")
        if errors:
            messagebox.showwarning("Trang in không hợp lệ", "\n".join(errors[:5]))
            return
        for i in editable:
            self.file_list[i].page_mode = "custom"
            self.file_list[i].page_range_text = raw
        self._refresh_tree(preserve_selection=True)
        self._update_stats_pill()
        self._log(f"Đã đặt trang in tùy chọn [{raw}] cho {len(editable)} tệp tin")
        if skipped:
            self._log(f"⚠️ Bỏ qua {len(skipped)} tệp đang in / đã in xong")
        self._save_config()

    def _apply_pages_to_all(self):
        """Apply current panel page setting to every editable file (spec §9)."""
        mode = PAGE_MODE_BY_LABEL.get(self._page_range_var.get(), "all")
        raw = self._custom_pages_var.get().strip()
        if mode == "custom":
            if not raw:
                messagebox.showinfo("Áp dụng cho tất cả", "Vui lòng nhập phạm vi trang tùy chọn trước.")
                return
            # Validate mẫu trên file đầu để báo lỗi sớm (số trang mỗi file check lúc in)
            try:
                parse_page_range(raw, 10 ** 9)
            except ValueError as exc:
                messagebox.showwarning("Trang in không hợp lệ", str(exc))
                return
        updated, skipped = 0, 0
        for f in self.file_list:
            if is_edit_locked(f.status):
                skipped += 1
                continue
            f.page_mode = mode
            f.page_range_text = raw if mode == "custom" else ""
            updated += 1
        self._refresh_tree(preserve_selection=True)
        self._update_stats_pill()
        msg = f"Đã áp dụng trang in [{self._page_range_var.get()}{' ' + raw if mode == 'custom' else ''}] cho {updated} tệp tin"
        if skipped:
            msg += f" (bỏ qua {skipped} tệp đang in/đã in)"
        self._log(msg)
        self._save_config()

    # ── Custom paper size ────────────────────────────────────────────

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

    def start_print(self, target_indices: Optional[list[int]] = None, action_title: str = "XÁC NHẬN BẮT ĐẦU IN"):
        if not self.file_list:
            messagebox.showwarning("Cảnh báo", "Hàng đợi trống — chưa có file nào.")
            return

        if self._is_printing:
            messagebox.showwarning("Cảnh báo", "Quá trình in đang diễn ra. Vui lòng chờ hoàn thành hoặc bấm Hủy in trước.")
            return

        # Determine target printers
        primary_printer = self._printer_var.get()
        if self._parallel_printers and len(self._parallel_printers) >= 2:
            printers_to_use = self._parallel_printers
        elif primary_printer and not primary_printer.startswith("("):
            printers_to_use = [primary_printer]
        else:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn máy in trước khi bắt đầu.")
            return

        failover_p = self._failover_printer_var.get()
        if failover_p in ("", "(Không dùng)"):
            failover_p = None

        orient_name = self._orient_var.get()
        duplex_name = self._duplex_var.get()
        duplex_val = DUPLEX_MODES.get(duplex_name, win32con.DMDUP_SIMPLEX)
        paper = self._resolve_gui_paper()
        fit = self._fit_to_page.get()
        binding_margin_mm = float(BINDING_MARGIN_OPTIONS.get(self._binding_margin_var.get(), 0))
        reverse_order = self._reverse_order_var.get()
        remove_blanks = self._remove_blanks_var.get()
        use_separator = self._separator_sheet_var.get()

        active_indices = target_indices if target_indices is not None else list(range(len(self.file_list)))
        if not active_indices:
            messagebox.showinfo("Thông báo", "Không có tệp tin nào để in.")
            return

        # Chạy chuẩn bị file (ensure_pdf/convert Office) dưới nền để không treo UI.
        prep_win = ctk.CTkToplevel(self)
        prep_win.title("Đang chuẩn bị tệp in...")
        prep_win.geometry("380x130")
        prep_win.transient(self)
        prep_win.grab_set()
        prep_win.resizable(False, False)
        ctk.CTkLabel(
            prep_win, text="⏳ Đang chuẩn bị tệp in (chuyển đổi Office → PDF)...",
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
        ).pack(padx=16, pady=(18, 8))
        prep_lbl = ctk.CTkLabel(prep_win, text=f"0/{len(active_indices)}", font=ctk.CTkFont(size=11))
        prep_lbl.pack(padx=16, pady=(0, 6))
        prep_bar = ctk.CTkProgressBar(prep_win, width=320)
        prep_bar.pack(padx=16, pady=(0, 14))
        prep_bar.set(0)
        self.configure(cursor="wait")

        # Snapshot cài đặt để thread nền không chạm widget Tk.
        # Trang in là thuộc tính RIÊNG từng file (finfo.page_mode/value).
        # Duplex hiệu dụng: override người dùng > auto hóa đơn > mặc định chung.
        _printers_snapshot = list(printers_to_use)
        _invoice_mode = self._invoice_mode_var.get()
        _duplex_file_map = {
            "simplex": win32con.DMDUP_SIMPLEX,
            "long": win32con.DMDUP_VERTICAL,
            "short": win32con.DMDUP_HORIZONTAL,
        }

        import threading

        def _bg_prepare():
            jobs: list[PrintJob] = []
            total_pages = 0
            total_copies_pages = 0
            total_blank_skipped = 0
            blank_logs: list[str] = []
            warn_suggest: list[int] = []  # file idx hóa đơn chờ hỏi (chế độ cảnh báo)
            auto_duplex_files: list[str] = []
            saved_sheets = 0
            error: Optional[str] = None

            for doc_num, i in enumerate(active_indices, 1):
                if i < 0 or i >= len(self.file_list):
                    continue
                finfo = self.file_list[i]
                try:
                    self.pdf_manager.ensure_pdf(finfo)
                except Exception as conv_exc:
                    error = f"Không thể chuẩn bị file: {finfo.filename}\n{conv_exc}"
                    break

                # Nhận diện bù cho file Office/scan (PDF đã làm lúc nạp)
                if _invoice_mode != INVOICE_OFF and not finfo.invoice_analyzed:
                    try:
                        self._analyze_invoice(finfo)
                    except Exception:
                        pass

                # 1. Trang in RIÊNG từng file (không dùng biến chung)
                try:
                    p0 = resolve_page_selection(finfo.page_mode, finfo.page_range_text, finfo.page_count)
                except ValueError as exc:
                    error = f"File: {finfo.filename}\n⚠ {exc}"
                    break

                if remove_blanks:
                    p0, skipped = PDFManager.filter_pages(finfo.pdf_path, p0, remove_blanks=True)
                    if skipped:
                        total_blank_skipped += len(skipped)
                        blank_logs.append(f"[{finfo.filename}] Đã tự động bỏ qua {len(skipped)} trang trắng")

                if not p0:
                    p0 = [0]

                # 2. Duplex hiệu dụng từng file (§6, §11-12, §18)
                file_duplex = duplex_val
                if finfo.duplex_override:
                    if finfo.duplex_mode in _duplex_file_map:
                        file_duplex = _duplex_file_map[finfo.duplex_mode]
                elif finfo.duplex_mode in _duplex_file_map:
                    file_duplex = _duplex_file_map[finfo.duplex_mode]
                    if finfo.duplex_auto:
                        auto_duplex_files.append(finfo.filename)
                        saved_sheets += (len(p0) - (len(p0) + 1) // 2) * finfo.copies
                elif (_invoice_mode == INVOICE_AUTO and finfo.invoice_detected
                        and len(p0) >= 2):
                    file_duplex = win32con.DMDUP_VERTICAL
                    finfo.duplex_mode = "long"
                    finfo.duplex_auto = True
                    auto_duplex_files.append(finfo.filename)
                    saved_sheets += (len(p0) - (len(p0) + 1) // 2) * finfo.copies
                elif (_invoice_mode == INVOICE_WARN and finfo.invoice_detected
                        and len(p0) >= 2):
                    warn_suggest.append(i)

                if orient_name == ORIENT_AUTO:
                    ori = (
                        win32con.DMORIENT_LANDSCAPE
                        if self.pdf_manager.is_landscape(finfo.pdf_path)
                        else win32con.DMORIENT_PORTRAIT
                    )
                else:
                    ori = ORIENTATIONS.get(orient_name, win32con.DMORIENT_PORTRAIT)

                if use_separator and len(active_indices) > 1:
                    try:
                        sep_pdf = PDFManager.create_separator_sheet(
                            filename=finfo.filename,
                            page_count=len(p0),
                            copies=finfo.copies,
                            printer_name=", ".join(_printers_snapshot),
                            doc_index=doc_num,
                        )
                        jobs.append(PrintJob(
                            index=-1, pdf_path=sep_pdf,
                            filename=f"Bìa phân cách #{doc_num}",
                            pages=[0], copies=1, paper_size=paper,
                            orientation=ori, duplex=win32con.DMDUP_SIMPLEX,
                            fit_to_page=True, binding_margin_mm=0.0,
                            reverse_order=False, is_separator=True,
                        ))
                    except Exception as sep_exc:
                        print(f"Error creating separator sheet: {sep_exc}")

                jobs.append(PrintJob(
                    index=i, pdf_path=finfo.pdf_path, filename=finfo.filename,
                    pages=p0, copies=finfo.copies, paper_size=paper,
                    orientation=ori, duplex=file_duplex, fit_to_page=fit,
                    binding_margin_mm=binding_margin_mm,
                    reverse_order=reverse_order, is_separator=False,
                ))
                total_pages += len(p0)
                total_copies_pages += len(p0) * finfo.copies
                self._async_queue.put((lambda d=doc_num, t=len(active_indices): (
                    prep_lbl.configure(text=f"{d}/{t}"),
                    prep_bar.set(d / t if t else 1),
                ), ()))

            def _on_prepared():
                nonlocal saved_sheets
                try:
                    prep_win.destroy()
                except Exception:
                    pass
                self.configure(cursor="")
                if error:
                    messagebox.showerror("Lỗi chuẩn bị tệp", error)
                    return
                for line in blank_logs:
                    self._log(line)
                # Chế độ "Chỉ cảnh báo": hỏi 1 lần duy nhất cho cả đợt (§7, §16)
                if warn_suggest and _invoice_mode == INVOICE_WARN:
                    names = [self.file_list[i].filename for i in warn_suggest
                             if 0 <= i < len(self.file_list)]
                    if names and messagebox.askyesno(
                        "Phát hiện hóa đơn điện tử",
                        f"💡 Phát hiện {len(names)} hóa đơn từ 2 trang thực tế trở lên:\n"
                        + "\n".join(f"• {n}" for n in names[:8])
                        + ("\n..." if len(names) > 8 else "")
                        + "\n\nHệ thống đề xuất: In 2 mặt – Cạnh dài.\nÁp dụng?",
                    ):
                        for i in warn_suggest:
                            if 0 <= i < len(self.file_list):
                                f = self.file_list[i]
                                f.duplex_mode = "long"
                                f.duplex_auto = True
                                auto_duplex_files.append(f.filename)
                        for job in jobs:
                            if job.index in warn_suggest and not job.is_separator:
                                job.duplex = win32con.DMDUP_VERTICAL
                        # Tính lại giấy tiết kiệm cho các file vừa áp dụng
                        for i in warn_suggest:
                            if 0 <= i < len(self.file_list):
                                f = self.file_list[i]
                                try:
                                    n = len(f.resolve_pages())
                                    saved_sheets += (n - (n + 1) // 2) * f.copies
                                except Exception:
                                    pass
                        self._refresh_tree(preserve_selection=True)
                    else:
                        for i in warn_suggest:
                            if 0 <= i < len(self.file_list):
                                self.file_list[i].duplex_override = True
                for name in auto_duplex_files:
                    self._log(f"🧾 {name} — tự động In 2 mặt (Cạnh dài)")
                self._finish_start_print(
                    jobs, active_indices, printers_to_use, failover_p,
                    duplex_name, paper, orient_name, reverse_order,
                    remove_blanks, use_separator,
                    total_pages, total_copies_pages, total_blank_skipped,
                    action_title, auto_duplex_files, saved_sheets,
                )

            self._async_queue.put((_on_prepared, ()))

        threading.Thread(target=_bg_prepare, daemon=True).start()
        return

    def _finish_start_print(self, jobs: list[PrintJob], active_indices: list[int],
                            printers_to_use: list[str], failover_p: Optional[str],
                            duplex_name: str, paper: str, orient_name: str,
                            reverse_order: bool, remove_blanks: bool,
                            use_separator: bool, total_pages: int,
                            total_copies_pages: int, total_blank_skipped: int,
                            action_title: str, auto_duplex_files: Optional[list] = None,
                            saved_sheets: int = 0):

        # Confirmation Dialog
        p_desc = f"{len(printers_to_use)} máy in ({', '.join(printers_to_use)})" if len(printers_to_use) > 1 else printers_to_use[0]
        f_desc = failover_p if failover_p else "Không kích hoạt"
        blank_desc = f"Bật (đã loại bỏ {total_blank_skipped} trang trắng)" if remove_blanks else "Tắt"
        rev_desc = "Trang cuối về trang 1" if reverse_order else "Trang 1 đến trang cuối"
        sep_desc = "Có chèn tờ bìa phân cách" if use_separator else "Không chèn"
        auto_duplex_files = auto_duplex_files or []
        if auto_duplex_files:
            inv_desc = (f"{len(auto_duplex_files)} hóa đơn tự 2 mặt "
                        f"(tiết kiệm ~{saved_sheets} tờ)")
        else:
            inv_desc = "Không có"

        msg = (
            f"• Số tệp tin:             {len(active_indices)}\n"
            f"• Tổng số trang in:       {total_pages}\n"
            f"• Tổng trang × bản:        {total_copies_pages}\n"
            f"• Máy in thực hiện:       {p_desc}\n"
            f"• Máy in dự phòng:        {f_desc}\n"
            f"• Kiểu in 2 mặt:          {duplex_name}\n"
            f"• Khổ giấy / Chiều:       {paper} / {orient_name}\n"
            f"• Lề đóng gáy sách:       {self._binding_margin_var.get()}\n"
            f"• Thứ tự trang:           {rev_desc}\n"
            f"• Bỏ trang trắng:         {blank_desc}\n"
            f"• Trang bìa phân cách:    {sep_desc}\n"
            f"• 🧾 Hóa đơn 2 mặt:       {inv_desc}\n"
        )
        if not messagebox.askyesno(action_title, msg, icon="question"):
            return

        # Auto save settings
        self._save_config()

        # Reset status for active target files
        for i in active_indices:
            if 0 <= i < len(self.file_list):
                self.file_list[i].status = FileStatus.WAITING
        self._refresh_tree(preserve_selection=True)
        self._reset_progress()
        self._set_printing_state(True)

        self._log(f"Bắt đầu in {len(active_indices)} tệp trên [{p_desc}]")
        self._log(f"Thiết lập: {duplex_name} | {paper} | {orient_name} | Gáy: {self._binding_margin_var.get()}")

        self.print_worker = MultiPrinterCoordinator(
            jobs=jobs,
            printer_names=printers_to_use,
            failover_printer=failover_p,
            on_file_start=self._cb_file_start,
            on_page_progress=self._cb_page_progress,
            on_file_complete=self._cb_file_complete,
            on_all_complete=self._cb_all_complete,
            on_error=self._cb_error,
            on_failover=self._cb_failover,
        )
        self.print_worker.start()

    def start_print_selected(self):
        """Prints only the currently selected items in the queue."""
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(
                "Chỉ in mục chọn",
                "Vui lòng chọn các tệp tin trong danh sách bạn muốn in.",
            )
            return

        indices = []
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    indices.append(idx)
            except Exception:
                pass

        if not indices:
            return

        if any(self.file_list[i].status == FileStatus.PRINTING for i in indices):
            messagebox.showwarning("Cảnh báo", "Có tệp tin trong danh sách chọn đang được in.")
            return

        self.start_print(target_indices=indices, action_title="XÁC NHẬN IN CÁC MỤC ĐÃ CHỌN")

    def retry_failed_prints(self):
        """Finds and reprints all files that failed or were cancelled."""
        indices = retry_indices(self.file_list)
        if not indices:
            messagebox.showinfo(
                "In lại tệp lỗi",
                "Không có tệp tin nào bị lỗi hoặc bị hủy trong hàng đợi.",
            )
            return

        self.start_print(target_indices=indices, action_title="XÁC NHẬN IN LẠI CÁC TỆP LỖI / ĐÃ HỦY")

    def _reset_selected_to_waiting(self):
        """Resets selected items' status back to Waiting (Chờ in)."""
        sel = self.tree.selection()
        if not sel:
            return
        count = 0
        for item in sel:
            try:
                idx = int(self.tree.item(item, "values")[0]) - 1
                if 0 <= idx < len(self.file_list):
                    if self.file_list[idx].status != FileStatus.PRINTING:
                        self.file_list[idx].status = FileStatus.WAITING
                        count += 1
            except Exception:
                pass
        if count:
            self._refresh_tree(preserve_selection=True)
            self._log(f"Đã đặt lại trạng thái 'Chờ in' cho {count} tệp tin")

    def pause_print(self):
        if self.print_worker is None:
            return
        if self.print_worker.is_paused:
            self.print_worker.resume()
            self.btn_pause.configure(text="⏸️ Tạm Dừng")
            self.status_lbl.configure(text="⏳ Đang in...")
            self._log("Tiếp tục in")
        else:
            self.print_worker.pause()
            self.btn_pause.configure(text="▶️ Tiếp Tục")
            self.status_lbl.configure(text="⏸️ Đang tạm dừng...")
            self._log("Tạm dừng in")

    def cancel_print(self):
        if self.print_worker is None:
            return
        if not messagebox.askyesno("Hủy in", "Bạn có chắc muốn dừng và hủy toàn bộ quá trình in?"):
            return
        self.print_worker.cancel()
        self._log("Đã gửi lệnh hủy in")

    # ── Worker Callbacks (Thread-Safe via Queue) ─────────────────────

    def _cb_file_start(self, idx, fn, printer_name=""):
        self._async_queue.put((self._on_file_start, (idx, fn, printer_name)))

    def _cb_page_progress(self, idx, fn, cur, tot, done, ftot):
        self._async_queue.put((self._on_page_progress, (idx, fn, cur, tot, done, ftot)))

    def _cb_file_complete(self, idx, fn, st, printer_name=""):
        self._async_queue.put((self._on_file_complete, (idx, fn, st, printer_name)))

    def _cb_all_complete(self, cancelled):
        self._async_queue.put((self._on_all_complete, (cancelled,)))

    def _cb_error(self, idx, fn, msg):
        self._async_queue.put((self._on_error, (idx, fn, msg)))

    def _cb_failover(self, fn, from_p, to_p, err):
        self._async_queue.put((self._on_failover, (fn, from_p, to_p, err)))

    # ── GUI Handlers ─────────────────────────────────────────────────

    def _on_file_start(self, index, filename, printer_name=""):
        p_tag = f" [{printer_name}]" if printer_name else ""
        if 0 <= index < len(self.file_list):
            self.file_list[index].status = FileStatus.PRINTING
            self._refresh_tree(preserve_selection=True)
            self.status_lbl.configure(text=f"⏳ Đang in: {filename}{p_tag}")
            self._log(f"{filename} — bắt đầu in ({self.file_list[index].copies} bản){p_tag}")
        elif index == -1:
            self.status_lbl.configure(text=f"📄 Đang in tờ bìa phân cách{p_tag}")
            self._log(f"📄 Bìa phân cách — bắt đầu in{p_tag}")
        self.file_progress.set(0)
        self.file_progress_lbl.configure(text="0 trang")

    def _on_page_progress(self, index, filename, cur, total, done, ftotal):
        self.file_progress.set(cur / total if total else 0)
        self.file_progress_lbl.configure(text=f"{cur}/{total} trang")
        prog = (done + cur / total) / ftotal if ftotal else 0
        self.total_progress.set(prog)
        self.total_progress_lbl.configure(text=f"{done}/{ftotal} mục")

    def _on_file_complete(self, index, filename, status, printer_name=""):
        p_tag = f" [{printer_name}]" if printer_name else ""
        if 0 <= index < len(self.file_list):
            self.file_list[index].status = status
            self._refresh_tree(preserve_selection=True)
        icon = STATUS_ICONS.get(status, "")
        self._log(f"{filename} — {icon} {status}{p_tag}")

    def _on_failover(self, filename, from_p, to_p, err):
        self._log(f"🛡️ TỰ ĐỘNG CHUYỂN DỰ PHÒNG: [{filename}] gặp sự cố trên '{from_p}' ({err}) → Đã tự động chuyển sang máy in '{to_p}'!")
        self.status_lbl.configure(text=f"🛡️ Chuyển dự phòng: {filename} → {to_p}")

    def _on_all_complete(self, was_cancelled):
        self._set_printing_state(False)
        if was_cancelled:
            for f in self.file_list:
                if f.status == FileStatus.WAITING:
                    f.status = FileStatus.CANCELLED
            self._refresh_tree(preserve_selection=True)
            self.status_lbl.configure(text="⏹️ Đã hủy quá trình in")
            self._log("Quá trình in đã bị hủy")
        else:
            self.total_progress.set(1.0)
            self.status_lbl.configure(text="✓ Hoàn thành tất cả tệp tin!")
            self._log("✓ Hoàn thành in tất cả tệp tin")

    def _on_error(self, index, filename, error_msg):
        self._log(f"⚠ LỖI [{filename}]: {error_msg}")

    # ── Progress Helpers ─────────────────────────────────────────────

    def _reset_progress(self):
        self.file_progress.set(0)
        self.file_progress_lbl.configure(text="0/0 trang")
        self.total_progress.set(0)
        self.total_progress_lbl.configure(text="0/0 tệp")
        self.status_lbl.configure(text="⏳ Đang khởi tạo...")

    def _set_printing_state(self, printing: bool):
        self._is_printing = printing
        st = "disabled" if printing else "normal"
        self.btn_start.configure(state=st)
        self.btn_print_selected.configure(state=st)
        self.btn_retry_failed.configure(state=st)
        self.btn_pause.configure(state="normal" if printing else "disabled")
        self.btn_cancel.configure(state="normal" if printing else "disabled")
        if printing:
            self.btn_pause.configure(text="⏸️ Tạm Dừng")
        else:
            self.print_worker = None

    # ── Logging ──────────────────────────────────────────────────────

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

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


class UpdaterMixin:

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

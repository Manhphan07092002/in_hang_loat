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


class PrintersMixin:

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

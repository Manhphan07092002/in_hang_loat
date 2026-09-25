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


class InvoiceMixin:

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

        self._build_sec_extra(card, _section, _row)

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

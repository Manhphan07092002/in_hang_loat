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


class QueueSessionMixin:

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

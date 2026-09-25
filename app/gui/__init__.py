"""
Modern GUI for PDF Batch Printer Pro — built with CustomTkinter.

Package layout (mỗi module ≤ 300 dòng, ráp thành PDFBatchPrinterApp):
  shell             Khung app, __init__, cửa sổ, icon, đóng app
  app_state         Config + kéo-thả file
  header            Thanh header, theme, hướng dẫn
  updater_ui        Tự động cập nhật
  invoice_ui        Nhận diện hóa đơn
  panels            Dashboard 2 cột + nhật ký + log
  action_card       Cụm tiến trình & nút in
  queue_view        Bảng hàng đợi + theme bảng
  queue_tools       Toolbar reflow, tooltip, hover
  queue_add         Nạp file vào hàng đợi
  queue_manage      Xóa/sắp xếp/di chuyển/cập nhật bảng
  queue_session     Lưu/mở phiên JSON
  copies_edit       Sửa inline số bản
  copies_bulk       Menu chuột phải, đặt bản hàng loạt
  pages_edit        Trang in riêng từng file
  pages_popup       Popup sửa Trang In trên bảng
  settings_shell    Khung card cài đặt
  settings_printer  Mục máy in
  settings_pages    Mục tệp đang chọn
  settings_paper    Khổ giấy, chiều, duplex, lề, dự phòng
  settings_extra    Hóa đơn, tùy chọn, in song song
  printers          Máy in, dialog song song
  print_prepare     Chuẩn bị job in (nền)
  print_run         Chạy/hủy/tạm dừng + callbacks tiến trình
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


from .paths import get_resource_path  # re-export cho tương thích


# NOTE: get_app_data_dir / CONFIG_FILE are imported from app.config_store
# (single source of truth, %APPDATA%-based) and re-exported for compatibility.

# Set theme globally before widget creation
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

from .shell import AppShell
from .app_state import AppStateMixin
from .header import HeaderMixin
from .updater_ui import UpdaterMixin
from .invoice_ui import InvoiceMixin
from .panels import PanelsMixin
from .action_card import ActionCardMixin
from .queue_view import QueueViewMixin
from .queue_tools import QueueToolsMixin
from .queue_add import QueueAddMixin
from .queue_manage import QueueManageMixin
from .queue_session import QueueSessionMixin
from .copies_edit import CopiesEditMixin
from .copies_bulk import CopiesBulkMixin
from .pages_edit import PagesEditMixin
from .pages_popup import PagesPopupMixin
from .settings_shell import SettingsShellMixin
from .settings_printer import SettingsPrinterMixin
from .settings_pages import SettingsPagesMixin
from .settings_paper import SettingsPaperMixin
from .settings_extra import SettingsExtraMixin
from .printers import PrintersMixin
from .print_prepare import PrintPrepareMixin
from .print_run import PrintRunMixin


class PDFBatchPrinterApp(
    AppStateMixin, HeaderMixin, UpdaterMixin, InvoiceMixin,
    PanelsMixin, ActionCardMixin, QueueViewMixin, QueueToolsMixin,
    QueueAddMixin, QueueManageMixin, QueueSessionMixin,
    CopiesEditMixin, CopiesBulkMixin, PagesEditMixin, PagesPopupMixin,
    SettingsShellMixin, SettingsPrinterMixin, SettingsPagesMixin,
    SettingsPaperMixin, SettingsExtraMixin, PrintersMixin,
    PrintPrepareMixin, PrintRunMixin, AppShell,
):
    """Modern Desktop GUI for PDF Batch Printer Pro (ráp từ các mixin)."""

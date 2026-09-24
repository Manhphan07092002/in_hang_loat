"""
Settings and constants for PDF Batch Printer Pro.
"""

import win32con

# ── Application ──────────────────────────────────────────────────────────────
APP_NAME = "PDF Batch Printer Pro"
APP_VERSION = "1.0.0.1"
WINDOW_TITLE = "PDF Batch Printer Pro — In Hàng Loạt Chuyên Nghiệp"
DEFAULT_SIZE = (1300, 780)
MIN_SIZE = (860, 560)

# ── Supported file types ─────────────────────────────────────────────────────
SUPPORTED_EXTENSIONS = {
    ".pdf": "PDF",
    ".doc": "Word",
    ".docx": "Word",
    ".rtf": "Word",
    ".ppt": "PowerPoint",
    ".pptx": "PowerPoint",
    ".xls": "Excel",
    ".xlsx": "Excel",
    ".jpg": "Ảnh",
    ".jpeg": "Ảnh",
    ".png": "Ảnh",
    ".bmp": "Ảnh",
    ".tiff": "Ảnh",
    ".tif": "Ảnh",
    ".gif": "Ảnh",
    ".webp": "Ảnh",
}

FILE_DIALOG_TYPES = [
    (
        "Tất cả file hỗ trợ",
        "*.pdf *.doc *.docx *.rtf *.ppt *.pptx *.xls *.xlsx "
        "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.gif *.webp",
    ),
    ("PDF (*.pdf)", "*.pdf"),
    ("Word (*.doc, *.docx, *.rtf)", "*.doc *.docx *.rtf"),
    ("PowerPoint (*.ppt, *.pptx)", "*.ppt *.pptx"),
    ("Excel (*.xls, *.xlsx)", "*.xls *.xlsx"),
    ("Hình ảnh (*.jpg, *.png, ...)", "*.jpg *.jpeg *.png *.bmp *.tiff *.tif *.gif *.webp"),
    ("Tất cả tệp (*.*)", "*.*"),
]

# ── Paper sizes → Windows constants ──────────────────────────────────────────
PAPER_SIZES = {
    "A3": win32con.DMPAPER_A3,
    "A4": win32con.DMPAPER_A4,
    "A5": win32con.DMPAPER_A5,
    "Letter": win32con.DMPAPER_LETTER,
    "Legal": win32con.DMPAPER_LEGAL,
}

PAPER_DIMENSIONS_PT = {
    "A3": (842, 1190),
    "A4": (595, 842),
    "A5": (420, 595),
    "Letter": (612, 792),
    "Legal": (612, 1008),
}

# ── Duplex modes ─────────────────────────────────────────────────────────────
DUPLEX_MODES = {
    "1 mặt": win32con.DMDUP_SIMPLEX,
    "2 mặt (lật như sách)": win32con.DMDUP_VERTICAL,
    "2 mặt (lật như lịch)": win32con.DMDUP_HORIZONTAL,
}

# ── Orientations ─────────────────────────────────────────────────────────────
ORIENTATIONS = {
    "Dọc": win32con.DMORIENT_PORTRAIT,
    "Ngang": win32con.DMORIENT_LANDSCAPE,
}

# ── File status ──────────────────────────────────────────────────────────────

class FileStatus:
    WAITING = "Chờ in"
    PRINTING = "Đang in"
    PRINTED = "Đã in"
    ERROR = "Lỗi"
    CANCELLED = "Đã hủy"
    SKIPPED = "Đã bỏ qua"


STATUS_ICONS = {
    FileStatus.WAITING: "○",
    FileStatus.PRINTING: "⏳",
    FileStatus.PRINTED: "✓",
    FileStatus.ERROR: "✕",
    FileStatus.CANCELLED: "■",
    FileStatus.SKIPPED: "⊘",
}

STATUS_COLORS = {
    FileStatus.WAITING: ("#64748B", "#94A3B8"),
    FileStatus.PRINTING: ("#2563EB", "#38BDF8"),
    FileStatus.PRINTED: ("#16A34A", "#4ADE80"),
    FileStatus.ERROR: ("#DC2626", "#F87171"),
    FileStatus.CANCELLED: ("#6B7280", "#9CA3AF"),
    FileStatus.SKIPPED: ("#D97706", "#FBBF24"),
}

# ── Page-range options ───────────────────────────────────────────────────────
PAGE_RANGE_ALL = "Tất cả các trang"
PAGE_RANGE_CUSTOM = "Trang tùy chọn"
PAGE_RANGE_OPTIONS = [PAGE_RANGE_ALL, PAGE_RANGE_CUSTOM]

# ── Orientation options ──────────────────────────────────────────────────────
ORIENT_AUTO = "Tự động"
ORIENT_PORTRAIT = "Dọc"
ORIENT_LANDSCAPE = "Ngang"
ORIENTATION_OPTIONS = [ORIENT_AUTO, ORIENT_PORTRAIT, ORIENT_LANDSCAPE]

# ── Binding Margin Options (mm) ──────────────────────────────────────────────
BINDING_MARGIN_OPTIONS = {
    "0 mm (Chuẩn)": 0,
    "10 mm (1 cm)": 10,
    "15 mm (1.5 cm)": 15,
    "20 mm (2 cm)": 20,
    "25 mm (2.5 cm)": 25,
}

# ── Preview ──────────────────────────────────────────────────────────────────
PREVIEW_DPI = 150
PREVIEW_MAX_CACHE = 30

# ── Modern Theme Palette (Light, Dark) ───────────────────────────────────────
THEME_COLORS = {
    "bg": ("#F8FAFC", "#0F172A"),
    "card": ("#FFFFFF", "#1E293B"),
    "card_alt": ("#F1F5F9", "#1E293B"),
    "card_border": ("#E2E8F0", "#334155"),
    "header_bg": ("#FFFFFF", "#1E293B"),
    "text": ("#0F172A", "#F8FAFC"),
    "text_muted": ("#64748B", "#94A3B8"),
    "primary": ("#2563EB", "#3B82F6"),
    "primary_hover": ("#1D4ED8", "#2563EB"),
    "success": ("#10B981", "#059669"),
    "success_hover": ("#059669", "#047857"),
    "danger": ("#EF4444", "#DC2626"),
    "danger_hover": ("#DC2626", "#B91C1C"),
    "warning": ("#F59E0B", "#D97706"),
    "warning_hover": ("#D97706", "#B45309"),
    "btn_secondary": ("#E2E8F0", "#334155"),
    "btn_secondary_hover": ("#CBD5E1", "#475569"),
    "btn_secondary_text": ("#334155", "#E2E8F0"),
    "accent_pill": ("#EFF6FF", "#1E3A8A"),
    "accent_pill_text": ("#1D4ED8", "#93C5FD"),
}

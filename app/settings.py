"""
Settings and constants for PDF Batch Printer Pro.
"""

import win32con

# ── Application ──────────────────────────────────────────────────────────────
APP_NAME = "PDF Batch Printer Pro"
APP_VERSION = "1.0.0.2"  # NGUỒN DUY NHẤT — build script tự sinh version_info.txt & installer.iss từ đây
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

# ── Custom paper size ────────────────────────────────────────────────────
# Người dùng nhập "Rộng x Cao (mm)", vd "210x297". printer_manager chuyển
# thành DMPAPER_USER + PaperWidth/PaperLength (đơn vị 1/10 mm).
CUSTOM_PAPER_LABEL = "Tùy chỉnh (R×C mm)..."

import re as _re
_CUSTOM_PAPER_RE = _re.compile(r"^\s*(\d+(?:\.\d+)?)\s*[x×]\s*(\d+(?:\.\d+)?)\s*(mm)?\s*$", _re.IGNORECASE)


def parse_custom_paper(text: str) -> tuple[float, float]:
    """Parse '210x297' → (210.0, 297.0) mm. Raise ValueError nếu sai/không hợp lệ."""
    m = _CUSTOM_PAPER_RE.match(text or "")
    if not m:
        raise ValueError("Khổ giấy tùy chỉnh phải dạng RộngxCao, ví dụ 210x297 (mm).")
    w, h = float(m.group(1)), float(m.group(2))
    if not (50 <= w <= 1500 and 50 <= h <= 1500):
        raise ValueError("Rộng/Cao phải từ 50 đến 1500 mm.")
    return w, h


def resolve_paper(paper: str) -> tuple[str, object]:
    """Trả về ('standard', DMPAPER_*) hoặc ('custom', (w_mm, h_mm))."""
    if paper in PAPER_SIZES:
        return "standard", PAPER_SIZES[paper]
    return "custom", parse_custom_paper(paper)


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
    "primary": ("#1D4ED8", "#3B82F6"),
    "primary_hover": ("#1E40AF", "#2563EB"),
    "success": ("#047857", "#059669"),
    "success_hover": ("#065F46", "#047857"),
    "danger": ("#B91C1C", "#DC2626"),
    "danger_hover": ("#991B1B", "#B91C1C"),
    "warning": ("#B45309", "#D97706"),
    "warning_hover": ("#92400E", "#B45309"),
    "btn_secondary": ("#E2E8F0", "#334155"),
    "btn_secondary_hover": ("#CBD5E1", "#475569"),
    "btn_secondary_text": ("#334155", "#E2E8F0"),
    "accent_pill": ("#EFF6FF", "#1E3A8A"),
    "accent_pill_text": ("#1D4ED8", "#93C5FD"),
}

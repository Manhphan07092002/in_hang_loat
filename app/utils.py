"""
Utility functions for PDF Batch Printer.
"""

import re
from datetime import datetime


def parse_page_range(text: str, max_pages: int) -> list[int]:
    """
    Parse a page range string into a sorted list of unique 1-indexed page
    numbers.

    Supported formats:
        "1"          → [1]
        "1-5"        → [1, 2, 3, 4, 5]
        "1,3,5"      → [1, 3, 5]
        "1-5,8-12"   → [1, 2, 3, 4, 5, 8, 9, 10, 11, 12]
        "2,4,7-10"   → [2, 4, 7, 8, 9, 10]

    Args:
        text: Page range string.
        max_pages: Maximum valid page number.

    Returns:
        Sorted list of 1-indexed page numbers.

    Raises:
        ValueError: If the range is invalid.
    """
    if not text or not text.strip():
        raise ValueError("Phạm vi trang không được để trống.")

    text = text.strip()

    # Only digits, spaces, commas, and hyphens allowed
    if not re.match(r'^[\d\s,\-]+$', text):
        raise ValueError("Phạm vi trang chứa ký tự không hợp lệ.")

    pages: set[int] = set()
    parts = text.split(',')

    for part in parts:
        part = part.strip()
        if not part:
            continue

        if '-' in part:
            range_parts = part.split('-')
            if len(range_parts) != 2 or not range_parts[0].strip() or not range_parts[1].strip():
                raise ValueError(f"Phạm vi không hợp lệ: '{part}'")
            try:
                start = int(range_parts[0].strip())
                end = int(range_parts[1].strip())
            except ValueError:
                raise ValueError(f"Số trang không hợp lệ trong: '{part}'")
            if start <= 0 or end <= 0:
                raise ValueError("Số trang phải lớn hơn 0.")
            if start > end:
                raise ValueError(f"Phạm vi không hợp lệ: {start} > {end}")
            if end > max_pages:
                raise ValueError(
                    f"Trang {end} vượt quá số trang tài liệu ({max_pages})."
                )
            pages.update(range(start, end + 1))
        else:
            try:
                page = int(part)
            except ValueError:
                raise ValueError(f"Số trang không hợp lệ: '{part}'")
            if page <= 0:
                raise ValueError("Số trang phải lớn hơn 0.")
            if page > max_pages:
                raise ValueError(
                    f"Trang {page} vượt quá số trang tài liệu ({max_pages})."
                )
            pages.add(page)

    if not pages:
        raise ValueError("Không có trang nào được chọn.")

    return sorted(pages)


def format_file_size(size_bytes: int) -> str:
    """Return a human-readable file size string."""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"


def get_timestamp() -> str:
    """Return the current time as ``HH:MM:SS``."""
    return datetime.now().strftime("%H:%M:%S")

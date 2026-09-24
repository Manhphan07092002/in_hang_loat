# -*- coding: utf-8 -*-
"""Pure queue logic for the print queue (no Tkinter dependency).

Extracted from the GUI God-class so it can be unit-tested headless and
reused by the CLI. All functions are pure: they never touch widgets.
"""
from __future__ import annotations

import os
from typing import Optional

from app.settings import FileStatus


def dedupe_paths(paths: list[str], existing_abs: set[str]) -> list[str]:
    """Filter *paths* to unseen absolute paths, updating *existing_abs*."""
    valid: list[str] = []
    for p in paths:
        ap = os.path.abspath(p)
        if ap in existing_abs:
            continue
        existing_abs.add(ap)
        valid.append(p)
    return valid


def is_edit_locked(status: str) -> bool:
    """Files being printed or already printed cannot have copies edited."""
    return status in (FileStatus.PRINTING, FileStatus.PRINTED)


def move_indices(n: int, indices: list[int], direction: int) -> set[int]:
    """Compute new index set after moving *indices* by *direction* (±1).

    *n* is the queue length. Returns the new selected-index set.
    Pure helper — caller swaps items in its own list identically.
    """
    if direction not in (1, -1):
        raise ValueError("direction must be +1 or -1")
    sel = sorted(i for i in indices if 0 <= i < n)
    if not sel:
        return set()
    new_sel = set(sel)
    ordered = sel if direction == -1 else list(reversed(sel))
    for i in ordered:
        j = i + direction
        if 0 <= j < n and j not in new_sel:
            new_sel.remove(i)
            new_sel.add(j)
    return new_sel


def apply_move(items: list, indices: list[int], direction: int) -> set[int]:
    """Swap items in *items* in place; return new selected-index set."""
    new_sel = move_indices(len(items), indices, direction)
    ordered = sorted(indices) if direction == -1 else sorted(indices, reverse=True)
    moved = set(indices)
    for i in ordered:
        j = i + direction
        if 0 <= j < len(items) and j not in moved:
            items[i], items[j] = items[j], items[i]
            moved.remove(i)
            moved.add(j)
    return new_sel


def queue_summary(file_list: list) -> tuple[int, int, Optional[int]]:
    """Return (total_files, total_copies, total_pages_or_None).

    total_pages is None when some page_count is not yet known.
    """
    total_files = len(file_list)
    total_copies = 0
    total_pages = 0
    calculating = False
    for f in file_list:
        try:
            c = int(f.copies)
            total_copies += c
            total_pages += int(f.page_count) * c
        except (ValueError, TypeError):
            calculating = True
    return total_files, total_copies, (None if calculating and total_files else total_pages)


def retry_indices(file_list: list) -> list[int]:
    """Indices of failed/cancelled files for the retry-failed action."""
    return [i for i, f in enumerate(file_list)
            if f.status in (FileStatus.ERROR, FileStatus.CANCELLED)]

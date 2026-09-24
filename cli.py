"""
CLI Interface for PDF Batch Printer Pro.

Supports two modes:
  1. Interactive CLI:
         python cli.py
  2. One-line Command Mode:
         python cli.py -f doc1.pdf doc2.docx -p "RICOH IM 2500" -c 2 --duplex long
         python cli.py --dir "C:/Documents" -r -p "HP LaserJet"
         python cli.py --list-printers
"""

import sys
import os
import argparse
import threading
import time

# Ensure UTF-8 on Windows console
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import win32con
from app.settings import (
    PAPER_SIZES, DUPLEX_MODES, ORIENTATIONS,
    SUPPORTED_EXTENSIONS, FileStatus,
)
from app.utils import parse_page_range, format_file_size
from app.pdf_manager import PDFManager, FileInfo
from app.printer_manager import PrinterManager
from app.print_worker import PrintJob, MultiPrinterCoordinator
from app.file_converter import cleanup_temp


# ═════════════════════════════════════════════════════════════════════════
#  CLI HELPER FUNCTIONS
# ═════════════════════════════════════════════════════════════════════════

def print_banner():
    print("=" * 68)
    print("  🖨️  PDF BATCH PRINTER PRO — Giao Diện Dòng Lệnh (CLI)")
    print("  Hỗ trợ: PDF • Word • Excel • PowerPoint • Hình ảnh")
    print("=" * 68)


def list_printers():
    """Print numbered list of available Windows printers."""
    printers = PrinterManager.get_printers()
    default = PrinterManager.get_default_printer()
    if not printers:
        print("❌ Không tìm thấy máy in nào trên hệ thống.")
        return []

    print("\n📋 DANH SÁCH MÁY IN KHẢ DỤNG:")
    print("-" * 68)
    for idx, p in enumerate(printers, 1):
        is_def = " (Mặc định)" if p == default else ""
        has_dup = " [Hỗ trợ 2 mặt]" if PrinterManager.supports_duplex(p) else ""
        print(f"  [{idx}] {p}{is_def}{has_dup}")
    print("-" * 68)
    return printers


def render_progress_bar(current: int, total: int, prefix: str = "", length: int = 30):
    """Draw a text-based progress bar."""
    percent = (current / total) * 100 if total else 0
    filled = int(length * current // total) if total else 0
    bar = "█" * filled + "░" * (length - filled)
    print(f"\r  {prefix} |{bar}| {current}/{total} ({percent:.0f}%)", end="", flush=True)


# ═════════════════════════════════════════════════════════════════════════
#  PRINT EXECUTION RUNNER
# ═════════════════════════════════════════════════════════════════════════

def execute_print_jobs(jobs: list[PrintJob], printer_name: str):
    """Run print jobs with real-time CLI progress reporting."""
    print(f"\n🚀 BẮT ĐẦU IN TRÊN [{printer_name}]...")
    print("=" * 68)

    all_done_event = threading.Event()

    total_files = len(jobs)

    def on_file_start(idx, filename, printer=""):
        tag = f" [{printer}]" if printer else ""
        print(f"\n📄 [{idx + 1}/{total_files}] Đang in: {filename} ({printer_name}{tag})")

    def on_page_progress(idx, filename, cur, tot, done_files, ftotal):
        render_progress_bar(cur, tot, prefix=f"Trang file #{idx + 1}")

    def on_file_complete(idx, filename, status, printer=""):
        print(f"\n  ✓ Hoàn thành file: {filename} [{status}]")

    def on_all_complete(cancelled):
        all_done_event.set()

    def on_error(idx, filename, error_msg):
        print(f"\n  ❌ LỖI file {filename}: {error_msg}")

    coordinator = MultiPrinterCoordinator(
        jobs=jobs,
        printer_names=[printer_name],
        on_file_start=on_file_start,
        on_page_progress=on_page_progress,
        on_file_complete=on_file_complete,
        on_all_complete=on_all_complete,
        on_error=on_error,
    )

    try:
        coordinator.start()
        while not all_done_event.is_set():
            time.sleep(0.2)
    except KeyboardInterrupt:
        print("\n\n⚠️ Người dùng bấm Ctrl+C — Đang hủy in...")
        coordinator.cancel()
        all_done_event.wait(timeout=5.0)

    print("\n" + "=" * 68)
    print("✨ QUÁ TRÌNH IN KẾT THÚC!")
    print("=" * 68)


# ═════════════════════════════════════════════════════════════════════════
#  INTERACTIVE CLI MODE
# ═════════════════════════════════════════════════════════════════════════

def run_interactive():
    print_banner()

    pdf_mgr = PDFManager()
    file_list: list[FileInfo] = []

    # ── Step 1: Add files ────────────────────────────────────────────
    print("\n[Bước 1/4] CHỌN TỆP TIN HOẶC THƯ MỤC CẦN IN:")
    print("  • Nhập đường dẫn tệp tin (PDF, Word, Excel, PPT, Ảnh) hoặc thư mục.")
    print("  • Có thể nhập nhiều file, cách nhau bởi dấu phẩy hoặc kéo thả vào đây.")
    print("  • Gõ 'xong' hoặc nhấn Enter khi đã thêm đủ file.\n")

    while True:
        try:
            user_input = input("👉 Đường dẫn file/thư mục (hoặc 'xong'): ").strip().strip('"').strip("'")
        except (EOFError, KeyboardInterrupt):
            print("\nĐã hủy.")
            return

        if not user_input or user_input.lower() in ("xong", "done", "ok", "q"):
            if file_list:
                break
            else:
                print("⚠️ Chưa có file nào trong hàng đợi. Vui lòng nhập ít nhất 1 file/thư mục.")
                continue

        # Check if input is a directory
        if os.path.isdir(user_input):
            rec_input = input("   Quét cả thư mục con? (c/k, mặc định k): ").strip().lower()
            rec = rec_input in ("c", "y", "yes", "co", "có")
            found = pdf_mgr.scan_folder(user_input, recursive=rec)
            if not found:
                print("   ❌ Không tìm thấy file hỗ trợ trong thư mục này.")
            else:
                print(f"   ⏳ Đang nạp {len(found)} tệp tin...")
                for p in found:
                    try:
                        info = pdf_mgr.add_file(p)
                        file_list.append(info)
                    except Exception as e:
                        print(f"   ⚠️ Bỏ qua {os.path.basename(p)}: {e}")
                print(f"   ✓ Đã thêm {len(found)} file (Tổng: {len(file_list)} file)")
        else:
            # Maybe comma-separated list of files
            paths = [p.strip().strip('"').strip("'") for p in user_input.split(",") if p.strip()]
            for p in paths:
                if not os.path.exists(p):
                    print(f"   ❌ Không tìm thấy đường dẫn: {p}")
                    continue
                try:
                    info = pdf_mgr.add_file(p)
                    file_list.append(info)
                    print(f"   ✓ Đã thêm: {info.filename} ({info.file_type}, {info.page_count} trang)")
                except Exception as e:
                    print(f"   ❌ Lỗi khi nạp {os.path.basename(p)}: {e}")

    # ── Display Summary Queue ────────────────────────────────────────
    print("\n📦 DANH SÁCH HÀNG ĐỢI IN:")
    print("-" * 68)
    print(f" {'#':<3} | {'Tên Tệp':<35} | {'Loại':<6} | {'Trang':<5} | {'Kích Thước'}")
    print("-" * 68)
    for i, f in enumerate(file_list, 1):
        fn = (f.filename[:32] + "...") if len(f.filename) > 35 else f.filename
        print(f" {i:<3} | {fn:<35} | {f.file_type:<6} | {f.page_count:<5} | {format_file_size(f.file_size)}")
    print("-" * 68)
    total_pgs = sum(f.page_count for f in file_list)
    print(f"  Tổng cộng: {len(file_list)} tệp tin, {total_pgs} trang.\n")

    # ── Step 2: Choose Printer ───────────────────────────────────────
    print("[Bước 2/4] CHỌN MÁY IN:")
    printers = list_printers()
    if not printers:
        return

    default_prn = PrinterManager.get_default_printer()
    def_idx = printers.index(default_prn) + 1 if default_prn in printers else 1

    prn_choice = input(f"👉 Chọn số máy in [1-{len(printers)}] (mặc định {def_idx}): ").strip()
    if not prn_choice:
        selected_printer = printers[def_idx - 1]
    else:
        try:
            choice_num = int(prn_choice)
            if 1 <= choice_num <= len(printers):
                selected_printer = printers[choice_num - 1]
            else:
                selected_printer = printers[def_idx - 1]
        except ValueError:
            selected_printer = printers[def_idx - 1]

    print(f"  ✓ Máy in đã chọn: {selected_printer}")

    # ── Step 3: Print Options ────────────────────────────────────────
    print("\n[Bước 3/4] CẤU HÌNH IN:")

    # Copies
    copies_input = input("👉 Số bản in cho mỗi file (mặc định 1): ").strip()
    copies = 1
    if copies_input.isdigit() and int(copies_input) > 0:
        copies = int(copies_input)

    # Duplex
    print("👉 Chọn chế độ in:")
    print("   [1] 1 mặt (Simplex)")
    print("   [2] 2 mặt - lật như sách (Long Edge)")
    print("   [3] 2 mặt - lật như lịch (Short Edge)")
    dup_choice = input("   Lựa chọn [1/2/3] (mặc định 1): ").strip()
    duplex_map = {
        "1": win32con.DMDUP_SIMPLEX,
        "2": win32con.DMDUP_VERTICAL,
        "3": win32con.DMDUP_HORIZONTAL,
    }
    duplex_val = duplex_map.get(dup_choice, win32con.DMDUP_SIMPLEX)

    # Paper Size
    print("👉 Khổ giấy: [1] A3  [2] A4  [3] A5  [4] Letter  [5] Legal  [6] Tùy chỉnh (RộngxCao mm)")
    paper_choice = input("   Lựa chọn [1-6] (mặc định A4, hoặc nhập trực tiếp vd 210x297): ").strip()
    paper_map = {"1": "A3", "2": "A4", "3": "A5", "4": "Letter", "5": "Legal"}
    if paper_choice in paper_map:
        paper_size = paper_map[paper_choice]
    elif not paper_choice:
        paper_size = "A4"
    elif paper_choice == "6":
        paper_size = input("   Nhập RộngxCao (mm), ví dụ 210x297: ").strip()
    else:
        paper_size = paper_choice  # cho phép nhập trực tiếp "210x297"
    from app.settings import resolve_paper as _resolve_paper
    try:
        _resolve_paper(paper_size)
    except ValueError as exc:
        print(f"❌ Khổ giấy không hợp lệ: {exc}")
        return

    # Orientation
    print("👉 Chiều in: [1] Tự động  [2] Dọc (Portrait)  [3] Ngang (Landscape)")
    ori_choice = input("   Lựa chọn [1/2/3] (mặc định Tự động): ").strip()
    ori_map = {"1": "Tự động", "2": "Dọc", "3": "Ngang"}
    orient_mode = ori_map.get(ori_choice, "Tự động")

    # Page range
    page_range_raw = input("👉 Phạm vi trang cần in (Enter để in tất cả, hoặc ví dụ 1-5, 8-10): ").strip()

    # ── Step 4: Build Jobs & Confirm ─────────────────────────────────
    jobs: list[PrintJob] = []
    for i, finfo in enumerate(file_list):
        try:
            pdf_mgr.ensure_pdf(finfo)
        except Exception as exc:
            print(f"❌ Không thể chuẩn bị file {finfo.filename}: {exc}")
            return
        if page_range_raw:
            try:
                p1 = parse_page_range(page_range_raw, finfo.page_count)
                p0 = [p - 1 for p in p1]
            except ValueError as exc:
                print(f"❌ Lỗi trang cho file {finfo.filename}: {exc}")
                return
        else:
            p0 = list(range(finfo.page_count))

        if orient_mode == "Tự động":
            ori = (
                win32con.DMORIENT_LANDSCAPE
                if pdf_mgr.is_landscape(finfo.pdf_path)
                else win32con.DMORIENT_PORTRAIT
            )
        else:
            ori = ORIENTATIONS.get(orient_mode, win32con.DMORIENT_PORTRAIT)

        jobs.append(PrintJob(
            index=i,
            pdf_path=finfo.pdf_path,
            filename=finfo.filename,
            pages=p0,
            copies=copies,
            paper_size=paper_size,
            orientation=ori,
            duplex=duplex_val,
            fit_to_page=True,
        ))

    print("\n[Bước 4/4] XÁC NHẬN IN:")
    print("-" * 68)
    print(f"  • Số tệp tin:        {len(jobs)}")
    print(f"  • Máy in:             {selected_printer}")
    print(f"  • Số bản:             {copies}")
    print(f"  • Khổ giấy:           {paper_size}")
    print(f"  • Chiều in:           {orient_mode}")
    print(f"  • Kiểu in:            {'1 mặt' if duplex_val == win32con.DMDUP_SIMPLEX else '2 mặt'}")
    print("-" * 68)

    confirm = input("👉 Bạn có muốn bắt đầu in ngay bây giờ? (C/K, mặc định C): ").strip().lower()
    if confirm in ("", "c", "y", "yes", "co", "có"):
        execute_print_jobs(jobs, selected_printer)
    else:
        print("Đã hủy in.")

    cleanup_temp()


# ═════════════════════════════════════════════════════════════════════════
#  COMMAND-LINE ARGUMENTS MODE
# ═════════════════════════════════════════════════════════════════════════

def run_cli_args(args):
    if args.list_printers:
        list_printers()
        return

    if getattr(args, "check_update", False):
        from app.updater import check_for_updates
        print_banner()
        print("⏳ Đang kiểm tra cập nhật...")
        info = check_for_updates()
        if info.get("error"):
            print(f"⚠️ {info['error']}")
        elif info.get("has_update"):
            print(f"🎉 Có bản mới v{info['latest']} (bạn đang dùng v{info['current']})")
            print(f"🔗 {info['url']}")
            if info.get("asset_name"):
                print(f"📦 File: {info['asset_name']}")
        else:
            print(f"✓ Bạn đang dùng bản mới nhất (v{info['current']}).")
        return

    pdf_mgr = PDFManager()
    collected_paths: list[str] = []

    # Collect from --files
    if args.files:
        for f in args.files:
            if os.path.exists(f):
                collected_paths.append(os.path.abspath(f))
            else:
                print(f"⚠️ Cảnh báo: Tệp không tồn tại: {f}")

    # Collect from --dir
    if args.dir:
        if os.path.isdir(args.dir):
            found = pdf_mgr.scan_folder(args.dir, recursive=args.recursive)
            collected_paths.extend(found)
        else:
            print(f"❌ Thư mục không tồn tại: {args.dir}")
            return

    if not collected_paths:
        print("❌ Không có tệp tin hợp lệ nào để in! Sử dụng -h để xem hướng dẫn.")
        return

    # Printer selection
    printer = args.printer
    if not printer:
        printer = PrinterManager.get_default_printer()
        if not printer:
            print("❌ Không tìm thấy máy in mặc định! Vui lòng chỉ định với -p \"Tên máy in\".")
            return

    # Duplex
    duplex_val = win32con.DMDUP_SIMPLEX
    if args.duplex:
        d = args.duplex.lower()
        if d in ("2", "long", "vertical", "sach", "book"):
            duplex_val = win32con.DMDUP_VERTICAL
        elif d in ("short", "horizontal", "lich", "calendar"):
            duplex_val = win32con.DMDUP_HORIZONTAL

    # Orientation
    orient_mode = args.orient.capitalize() if args.orient else "Tự động"
    from app.settings import resolve_paper as _resolve_paper2
    paper_size = args.paper.upper() if args.paper and "x" not in args.paper.lower() else (args.paper or "A4")
    try:
        _resolve_paper2(paper_size)
    except ValueError as exc:
        print(f"❌ Khổ giấy không hợp lệ (--paper): {exc}")
        return

    # Load files
    print_banner()
    print(f"⏳ Đang nạp {len(collected_paths)} tệp tin...")
    file_list: list[FileInfo] = []
    for p in collected_paths:
        try:
            info = pdf_mgr.add_file(p)
            file_list.append(info)
        except Exception as exc:
            print(f"⚠️ Bỏ qua {os.path.basename(p)}: {exc}")

    if not file_list:
        print("❌ Không thể nạp tệp tin nào.")
        return

    # Build jobs
    jobs: list[PrintJob] = []
    for i, finfo in enumerate(file_list):
        try:
            pdf_mgr.ensure_pdf(finfo)
        except Exception as exc:
            print(f"❌ Không thể chuẩn bị file {finfo.filename}: {exc}")
            return
        if args.pages and args.pages.lower() != "all":
            try:
                p1 = parse_page_range(args.pages, finfo.page_count)
                p0 = [p - 1 for p in p1]
            except ValueError as exc:
                print(f"❌ Lỗi trang cho file {finfo.filename}: {exc}")
                return
        else:
            p0 = list(range(finfo.page_count))

        if orient_mode == "Tự động":
            ori = (
                win32con.DMORIENT_LANDSCAPE
                if pdf_mgr.is_landscape(finfo.pdf_path)
                else win32con.DMORIENT_PORTRAIT
            )
        else:
            ori = ORIENTATIONS.get(orient_mode, win32con.DMORIENT_PORTRAIT)

        jobs.append(PrintJob(
            index=i,
            pdf_path=finfo.pdf_path,
            filename=finfo.filename,
            pages=p0,
            copies=max(1, args.copies),
            paper_size=paper_size,
            orientation=ori,
            duplex=duplex_val,
            fit_to_page=not args.no_fit,
        ))

    execute_print_jobs(jobs, printer)
    cleanup_temp()


# ═════════════════════════════════════════════════════════════════════════
#  MAIN ENTRY POINT
# ═════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="PDF Batch Printer Pro — In Hàng Loạt CLI",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "-f", "--files", nargs="+",
        help="Danh sách tệp tin cần in (PDF, Word, Excel, PPT, Ảnh)",
    )
    parser.add_argument(
        "-d", "--dir",
        help="Thư mục chứa các tệp tin cần in",
    )
    parser.add_argument(
        "-r", "--recursive", action="store_true",
        help="Quét cả thư mục con (khi dùng với -d / --dir)",
    )
    parser.add_argument(
        "-p", "--printer",
        help="Tên máy in Windows (mặc định lấy máy in Default)",
    )
    parser.add_argument(
        "-c", "--copies", type=int, default=1,
        help="Số bản in (mặc định: 1)",
    )
    parser.add_argument(
        "--pages",
        help="Phạm vi trang in, ví dụ: 1-5, 8-12 (mặc định: all)",
    )
    parser.add_argument(
        "--duplex", choices=["1", "simplex", "long", "vertical", "short", "horizontal"],
        default="1", help="In 2 mặt: '1' (1 mặt), 'long' (lật như sách), 'short' (lật như lịch)",
    )
    parser.add_argument(
        "--paper", default="A4",
        help="Khổ giấy: A3/A4/A5/Letter/Legal hoặc tùy chỉnh WxH (vd 210x297, mm)",
    )
    parser.add_argument(
        "--orient", choices=["auto", "portrait", "landscape"], default="auto",
        help="Chiều giấy in: auto (Tự động), portrait (Dọc), landscape (Ngang)",
    )
    parser.add_argument(
        "--no-fit", action="store_true",
        help="Không tự động scale vừa trang giấy",
    )
    parser.add_argument(
        "-l", "--list-printers", action="store_true",
        help="Liệt kê danh sách tất cả máy in hiện có và thoát",
    )
    parser.add_argument(
        "--check-update", action="store_true",
        help="Kiểm tra bản mới trên GitHub Releases rồi thoát",
    )

    args = parser.parse_args()

    # If no arguments are passed, launch interactive CLI mode
    if len(sys.argv) == 1:
        run_interactive()
    else:
        run_cli_args(args)


if __name__ == "__main__":
    main()

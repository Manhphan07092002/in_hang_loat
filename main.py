"""
PDF Batch Printer Pro — Main Entry Point (Supports both GUI and CLI)

Usage:
    python main.py              # Launch GUI
    python main.py --cli        # Launch Interactive CLI
    python main.py -h           # View CLI Options
"""

import sys
import os

# Ensure UTF-8 output on Windows console
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def main():
    # If any CLI arguments are passed (e.g. -f, --dir, --cli, --list-printers, -h)
    if len(sys.argv) > 1:
        if "--cli" in sys.argv:
            sys.argv.remove("--cli")
        from cli import main as cli_main
        cli_main()
        return

    # Default: Launch GUI mode
    print("=" * 60)
    print("  PDF BATCH PRINTER PRO — Dang khoi chay giao dien GUI...")
    print("  Tip: Ban co the chay 'python cli.py' de dung giao dien dong lenh.")
    print("=" * 60)

    try:
        from app.gui import PDFBatchPrinterApp

        app = PDFBatchPrinterApp()
        app.show_window()
        print("[OK] Cua so ung dung da san sang tren Desktop!")
        app.mainloop()
    except Exception as exc:
        print(f"\n[LOI] Khoi dong that bai: {exc}")
        import traceback
        traceback.print_exc()
        if sys.stdin and sys.stdin.isatty():
            try:
                input("\nNhan Enter de thoat...")
            except Exception:
                pass


if __name__ == "__main__":
    main()

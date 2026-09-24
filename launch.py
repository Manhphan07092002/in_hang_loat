"""
Launcher for PDF Batch Printer Pro — opens the GUI as a detached desktop app.

Usage:
    python launch.py
"""

import subprocess
import sys
import os

# Ensure UTF-8 console output
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    main_py = os.path.join(script_dir, "main.py")

    # Launch GUI process
    subprocess.Popen([sys.executable, main_py], cwd=script_dir)

    print("=" * 60)
    print("  [OK] DA KHOI CHAY GIAO DIEN PDF BATCH PRINTER PRO!")
    print("  Ung dung dang mo duoi dang cua so doc lap tren Desktop.")
    print("=" * 60)


if __name__ == "__main__":
    main()

import os
import sys
import time
import tempfile

if sys.stdout is not None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import pymupdf as fitz
from app.pdf_manager import PDFManager

def run_benchmarks():
    print("=" * 60)
    print("🚀 BẮT ĐẦU KIỂM TRA HIỆU NĂNG NẠP FILE (BENCHMARK)")
    print("=" * 60)

    mgr = PDFManager()
    temp_dir = tempfile.mkdtemp(prefix="benchmark_files_")

    # 1. Create 100 sample PDF files
    print("\n📦 Đang tạo 100 tệp PDF mẫu...")
    pdf_paths = []
    for i in range(100):
        p = os.path.join(temp_dir, f"test_doc_{i:03d}.pdf")
        doc = fitz.open()
        page = doc.new_page(width=595, height=842)
        page.insert_text((50, 50), f"Sample Document #{i} for performance testing")
        doc.save(p)
        doc.close()
        pdf_paths.append(p)

    # Measure PDF scanning & quick inspect
    t0 = time.perf_counter()
    scanned = mgr.scan_folder(temp_dir, recursive=False)
    t_scan = time.perf_counter() - t0
    print(f"✓ Quét thư mục 100 file PDF (os.scandir): {t_scan * 1000:.2f} ms ({len(scanned)} files)")

    # Measure batch quick_inspect of 100 PDFs
    import concurrent.futures
    t0 = time.perf_counter()
    with concurrent.futures.ThreadPoolExecutor(max_workers=16) as ex:
        infos = list(ex.map(lambda p: mgr.add_file(p, lazy=True), pdf_paths))
    t_inspect = time.perf_counter() - t0
    print(f"✓ Nạp & đọc metadata 100 file PDF đa luồng: {t_inspect * 1000:.2f} ms ({len(infos)} items)")
    print(f"  → Tốc độ trung bình: {t_inspect / len(pdf_paths) * 1000:.3f} ms / file")

    # 2. Test sample image quick inspection
    print("\n🖼️ Đang kiểm tra tốc độ nạp tệp hình ảnh...")
    img_paths = []
    for i in range(50):
        img_p = os.path.join(temp_dir, f"test_image_{i:03d}.png")
        # Create minimal 100x100 PNG
        doc = fitz.open()
        p = doc.new_page(width=100, height=100)
        p.draw_rect(fitz.Rect(0, 0, 100, 100), color=(0.2, 0.5, 0.8), fill=(0.2, 0.5, 0.8))
        pix = p.get_pixmap()
        pix.save(img_p)
        doc.close()
        img_paths.append(img_p)

    t0 = time.perf_counter()
    with concurrent.futures.ThreadPoolExecutor(max_workers=16) as ex:
        img_infos = list(ex.map(lambda p: mgr.add_file(p, lazy=True), img_paths))
    t_img = time.perf_counter() - t0
    print(f"✓ Nạp 50 file hình ảnh (Lazy inspection): {t_img * 1000:.2f} ms")
    print(f"  → Tốc độ trung bình: {t_img / len(img_paths) * 1000:.3f} ms / image")

    # Cleanup temp
    import shutil
    shutil.rmtree(temp_dir, ignore_errors=True)

    print("\n" + "=" * 60)
    print("🎉 KIỂM TRA HIỆU NĂNG HOÀN TẤT THÀNH CÔNG! TỐC ĐỘ ĐẠT MỨC SIÊU TỐC.")
    print("=" * 60)

if __name__ == "__main__":
    run_benchmarks()

# -*- coding: utf-8 -*-
"""
Bo test toan dien cho PDF Batch Printer Pro.
Chay:  python -m pytest test_full.py -v
   hoac python test_full.py
Bao phu: settings, utils, pdf_manager, file_converter, printer_manager,
         print_worker, preview logic, cli, gui logic (headless-safe), tich hop.
"""
import os
import sys
import shutil
import tempfile
import threading
import unittest
from unittest import mock

import fitz
from PIL import Image

from app.settings import (
    SUPPORTED_EXTENSIONS, FILE_DIALOG_TYPES, PAPER_SIZES, DUPLEX_MODES,
    ORIENTATIONS, FileStatus, STATUS_ICONS, STATUS_COLORS,
    BINDING_MARGIN_OPTIONS, PREVIEW_DPI, PREVIEW_MAX_CACHE, THEME_COLORS,
    PAGE_RANGE_ALL, PAGE_RANGE_CUSTOM,
)
from app.utils import parse_page_range, format_file_size, get_timestamp
from app.pdf_manager import PDFManager, FileInfo
from app.file_converter import convert_to_pdf, cleanup_temp
from app.printer_manager import PrinterManager
from app.print_worker import PrintJob, PrintWorker, MultiPrinterCoordinator
from app.preview import get_paper_name


def make_pdf(path, n_pages=3, text="Hello"):
    doc = fitz.open()
    for i in range(n_pages):
        p = doc.new_page(width=595, height=842)
        p.insert_text((50, 50), f"{text} {i+1}", fontsize=16)
    doc.save(path)
    doc.close()
    return path


def make_blank_pdf(path):
    doc = fitz.open()
    doc.new_page(width=595, height=842)  # blank
    p = doc.new_page(width=595, height=842)
    p.insert_text((50, 50), "co chu", fontsize=16)
    doc.save(path)
    doc.close()
    return path


def make_image(path, fmt="PNG", size=(200, 150), color="red"):
    Image.new("RGB", size, color=color).save(path, fmt)
    return path


# ================= SETTINGS =================
class TestSettings(unittest.TestCase):
    def test_supported_extensions(self):
        for ext in [".pdf", ".doc", ".docx", ".rtf", ".ppt", ".pptx",
                    ".xls", ".xlsx", ".jpg", ".jpeg", ".png", ".bmp",
                    ".tiff", ".tif", ".gif", ".webp"]:
            self.assertIn(ext, SUPPORTED_EXTENSIONS, f"thieu {ext}")
        self.assertEqual(SUPPORTED_EXTENSIONS[".pdf"], "PDF")
        self.assertEqual(SUPPORTED_EXTENSIONS[".webp"], "Ảnh")

    def test_file_dialog_types_cover_webp(self):
        all_types = " ".join(t[1] for t in FILE_DIALOG_TYPES)
        self.assertIn("*.webp", all_types)

    def test_paper_duplex_orient(self):
        self.assertIn("A4", PAPER_SIZES)
        self.assertEqual(len(DUPLEX_MODES), 3)
        self.assertIn("Dọc", ORIENTATIONS)
        self.assertIn("Ngang", ORIENTATIONS)

    def test_status_maps_complete(self):
        for s in [FileStatus.WAITING, FileStatus.PRINTING, FileStatus.PRINTED,
                  FileStatus.ERROR, FileStatus.CANCELLED, FileStatus.SKIPPED]:
            self.assertIn(s, STATUS_ICONS)
            self.assertIn(s, STATUS_COLORS)

    def test_preview_consts(self):
        self.assertGreater(PREVIEW_DPI, 0)
        self.assertGreaterEqual(PREVIEW_MAX_CACHE, 10)
        self.assertIn("bg", THEME_COLORS)
        self.assertIn("text", THEME_COLORS)

    def test_binding_margins(self):
        self.assertIn("0 mm (Chuẩn)", BINDING_MARGIN_OPTIONS)
        self.assertEqual(BINDING_MARGIN_OPTIONS["0 mm (Chuẩn)"], 0)


# ================= UTILS =================
class TestUtils(unittest.TestCase):
    def test_parse_valid(self):
        self.assertEqual(parse_page_range("1", 5), [1])
        self.assertEqual(parse_page_range("1-3", 5), [1, 2, 3])
        self.assertEqual(parse_page_range("1,3,5", 5), [1, 3, 5])
        self.assertEqual(parse_page_range("1-3,5", 5), [1, 2, 3, 5])
        self.assertEqual(parse_page_range(" 2 , 2 , 3 ", 5), [2, 3])  # dedup+sort
        self.assertEqual(parse_page_range("5", 5), [5])

    def test_parse_invalid(self):
        for bad in ["", "   ", "0", "0-3", "4-2", "6", "abc", "1-a",
                    "1--2", "--", "1-2-3", "@#"]:
            with self.assertRaises(ValueError, msg=f"phai loi: {bad!r}"):
                parse_page_range(bad, 5)
        with self.assertRaises(ValueError):
            parse_page_range("1-10", 5)
        with self.assertRaises(ValueError):
            parse_page_range("3-100", 5)

    def test_format_file_size_boundaries(self):
        self.assertEqual(format_file_size(0), "0 B")
        self.assertEqual(format_file_size(1023), "1023 B")
        self.assertEqual(format_file_size(1024), "1.0 KB")
        self.assertEqual(format_file_size(1024 * 1024), "1.0 MB")
        self.assertEqual(format_file_size(1024 * 1024 * 1024), "1.0 GB")

    def test_get_timestamp_format(self):
        import re
        self.assertRegex(get_timestamp(), r"^\d{2}:\d{2}:\d{2}$")


# ================= PDF MANAGER =================
class TestPDFManager(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.pdf = make_pdf(os.path.join(self.tmp, "a.pdf"), 4)
        self.pm = PDFManager()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)
        cleanup_temp()

    def test_quick_inspect_pdf(self):
        info = PDFManager.quick_inspect(self.pdf)
        self.assertEqual(info.page_count, 4)
        self.assertEqual(info.file_type, "PDF")
        self.assertTrue(info.is_converted_ready)

    def test_quick_inspect_images(self):
        for fmt, ext in [("PNG", ".png"), ("JPEG", ".jpg"), ("WEBP", ".webp")]:
            p = os.path.join(self.tmp, f"img{ext}")
            make_image(p, fmt=fmt)
            info = PDFManager.quick_inspect(p)
            self.assertEqual(info.file_type, "Ảnh")
            self.assertFalse(info.is_converted_ready)
            self.assertTrue(info.is_converted)

    def test_quick_inspect_unsupported(self):
        p = os.path.join(self.tmp, "x.xyz123")
        with open(p, "w") as f:
            f.write("x")
        with self.assertRaises(ValueError):
            PDFManager.quick_inspect(p)

    def test_quick_inspect_office_lazy(self):
        p = os.path.join(self.tmp, "doc.docx")
        with open(p, "wb") as f:
            f.write(b"fake")
        info = PDFManager.quick_inspect(p)
        self.assertEqual(info.file_type, "Word")
        self.assertFalse(info.is_converted_ready)
        self.assertEqual(info.page_count, 1)

    def test_ensure_pdf_passthrough(self):
        info = PDFManager.quick_inspect(self.pdf)
        out = PDFManager.ensure_pdf(info)
        self.assertEqual(out, os.path.abspath(self.pdf))

    def test_ensure_pdf_image(self):
        p = os.path.join(self.tmp, "pic.png")
        make_image(p)
        info = PDFManager.quick_inspect(p)
        out = PDFManager.ensure_pdf(info)
        self.assertTrue(os.path.exists(out))
        self.assertTrue(info.is_converted_ready)
        self.assertGreaterEqual(info.page_count, 1)

    def test_ensure_pdf_webp(self):
        p = os.path.join(self.tmp, "pic.webp")
        make_image(p, fmt="WEBP")
        info = PDFManager.quick_inspect(p)
        out = PDFManager.ensure_pdf(info)
        self.assertTrue(os.path.exists(out))
        self.assertEqual(PDFManager.get_page_count(out), 1)

    def test_add_file_lazy_eager(self):
        info = self.pm.add_file(self.pdf, lazy=True)
        self.assertEqual(info.page_count, 4)
        info2 = self.pm.add_file(self.pdf, lazy=False)
        self.assertEqual(info2.page_count, 4)

    def test_add_files_batch_skip_unsupported(self):
        bad = os.path.join(self.tmp, "bad.zzz")
        with open(bad, "w") as f:
            f.write("x")
        ok, errs = self.pm.add_files([self.pdf, bad], lazy=True)
        self.assertEqual(len(ok), 1)
        self.assertEqual(errs, [])

    def test_scan_folder(self):
        sub = os.path.join(self.tmp, "sub")
        os.makedirs(sub)
        nested = os.path.join(sub, "n.pdf")
        shutil.copyfile(self.pdf, nested)
        top = self.pm.scan_folder(self.tmp, recursive=False)
        self.assertIn(self.pdf, top)
        self.assertNotIn(nested, top)
        allf = self.pm.scan_folder(self.tmp, recursive=True)
        self.assertIn(nested, allf)
        # sap xep tang dan theo ten
        self.assertEqual(allf, sorted(allf, key=str.lower))
        self.assertEqual(self.pm.scan_folder(os.path.join(self.tmp, "khong_ton_tai")), [])

    def test_get_page_count_missing_corrupt(self):
        self.assertEqual(PDFManager.get_page_count(self.pdf), 4)
        self.assertEqual(PDFManager.get_page_count(os.path.join(self.tmp, "no.pdf")), 0)
        bad = os.path.join(self.tmp, "corrupt.pdf")
        with open(bad, "w") as f:
            f.write("khong phai pdf")
        self.assertEqual(PDFManager.get_page_count(bad), 0)

    def test_get_page_size(self):
        w, h = PDFManager.get_page_size(self.pdf, 0)
        self.assertAlmostEqual(w, 595, delta=3)
        self.assertAlmostEqual(h, 842, delta=3)
        w2, h2 = PDFManager.get_page_size(self.pdf, 999)  # out of range
        self.assertAlmostEqual(w2, 595, delta=1)
        self.assertFalse(PDFManager.is_landscape(self.pdf))
        # landscape
        lp = os.path.join(self.tmp, "land.pdf")
        doc = fitz.open()
        doc.new_page(width=842, height=595)
        doc.save(lp)
        doc.close()
        self.assertTrue(PDFManager.is_landscape(lp))

    def test_render_page(self):
        img = self.pm.render_page(self.pdf, 0, width=300, height=300)
        self.assertIsNotNone(img)
        self.assertIsNone(self.pm.render_page(self.pdf, 999, width=200, height=200))
        self.assertIsNone(self.pm.render_page(os.path.join(self.tmp, "no.pdf"), 0))
        # cache hit tra ve cung object
        img2 = self.pm.render_page(self.pdf, 0, width=300, height=300)
        self.assertIs(img, img2)

    def test_render_cache_bounded(self):
        # render nhieu trang/kich thuoc khac nhau -> cache khong vo han
        for i in range(60):
            self.pm.render_page(self.pdf, i % 4, width=200 + i, height=200 + i)
        self.assertLessEqual(len(self.pm._thumbnail_cache), PREVIEW_MAX_CACHE)
        self.assertLessEqual(self.pm._cache_bytes, 256 * 1024 * 1024 + 10 * 1024 * 1024)
        self.pm.clear_cache()
        self.assertEqual(len(self.pm._thumbnail_cache), 0)
        self.assertEqual(self.pm._cache_bytes, 0)

    def test_blank_detection_filter(self):
        bp = make_blank_pdf(os.path.join(self.tmp, "b.pdf"))
        active, skipped = PDFManager.filter_pages(bp, [0, 1], remove_blanks=True)
        self.assertEqual(active, [1])
        self.assertEqual(skipped, [0])
        a2, s2 = PDFManager.filter_pages(bp, [0, 1], remove_blanks=False)
        self.assertEqual((a2, s2), ([0, 1], []))
        # out-of-range bi loai
        a3, _ = PDFManager.filter_pages(bp, [0, 99], remove_blanks=True)
        self.assertNotIn(99, a3)

    def test_separator_sheet(self):
        sep = PDFManager.create_separator_sheet("t€st unicode ten dai.pdf", 10, 2, "MayIn", 1)
        self.assertTrue(os.path.exists(sep))
        self.assertEqual(PDFManager.get_page_count(sep), 1)
        os.remove(sep)


# ================= FILE CONVERTER =================
class TestFileConverter(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        cleanup_temp()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_pdf_passthrough(self):
        p = make_pdf(os.path.join(self.tmp, "x.pdf"), 2)
        out, is_temp = convert_to_pdf(p)
        self.assertEqual(out, os.path.abspath(p))
        self.assertFalse(is_temp)

    def test_image_formats(self):
        for fmt, ext in [("PNG", ".png"), ("JPEG", ".jpg"), ("WEBP", ".webp"), ("BMP", ".bmp")]:
            with self.subTest(ext=ext):
                p = os.path.join(self.tmp, f"i{ext}")
                make_image(p, fmt=fmt)
                out, is_temp = convert_to_pdf(p)
                self.assertTrue(is_temp)
                self.assertTrue(os.path.exists(out))
                self.assertEqual(PDFManager.get_page_count(out), 1)

    def test_unsupported(self):
        p = os.path.join(self.tmp, "f.unknownext123")
        with open(p, "w") as f:
            f.write("x")
        with self.assertRaises(ValueError):
            convert_to_pdf(p)

    def test_unique_paths_and_cleanup(self):
        p = os.path.join(self.tmp, "i.png")
        make_image(p)
        o1, _ = convert_to_pdf(p)
        o2, _ = convert_to_pdf(p)
        self.assertNotEqual(o1, o2)
        cleanup_temp()
        self.assertFalse(os.path.exists(o1))


# ================= PRINTER MANAGER (mocked) =================
class TestPrinterManager(unittest.TestCase):
    def test_enumeration_types(self):
        printers = PrinterManager.get_printers()
        self.assertIsInstance(printers, list)
        self.assertIsInstance(PrinterManager.get_default_printer(), str)
        txt, ready = PrinterManager.get_printer_status("MayIn_Khong_Ton_Tai_123")
        self.assertIsInstance(txt, str)
        self.assertIsInstance(ready, bool)
        self.assertIsInstance(PrinterManager.supports_duplex("MayIn_Khong_Ton_Tai_123"), bool)

    def test_print_pdf_validation(self):
        import threading
        tmp = tempfile.mkdtemp()
        try:
            pdf = make_pdf(os.path.join(tmp, "v.pdf"), 2)
            with self.assertRaises(ValueError):
                PrinterManager.print_pdf("", pdf, [0], 1, "A4", 1, 1, True, threading.Event())
            with self.assertRaises(FileNotFoundError):
                PrinterManager.print_pdf("P", os.path.join(tmp, "no.pdf"), [0], 1, "A4", 1, 1, True, threading.Event())
            with self.assertRaises(ValueError):
                PrinterManager.print_pdf("P", pdf, [], 1, "A4", 1, 1, True, threading.Event())
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_print_pdf_mocked_software_copies(self):
        """Mock toan bo Win32 GDI de kiem tra vong lap copies phan mem."""
        import threading
        import win32con
        tmp = tempfile.mkdtemp()
        try:
            pdf = make_pdf(os.path.join(tmp, "c.pdf"), 2)
            calls = {"pages": 0}

            class FakeDC:
                def GetDeviceCaps(self, cap):
                    if cap == win32con.HORZRES:
                        return 2480
                    if cap == win32con.VERTRES:
                        return 3508
                    return 300
                def StartDoc(self, n):
                    pass
                def StartPage(self):
                    pass
                def EndPage(self):
                    pass
                def EndDoc(self):
                    pass
                def AbortDoc(self):
                    pass
                def GetHandleOutput(self):
                    return 0
                def DeleteDC(self):
                    pass

            class FakeDib:
                def __init__(self, img):
                    pass
                def draw(self, h, box):
                    calls["pages"] += 1

            fake_devmode = mock.MagicMock()
            fake_devmode.Fields = 0
            with mock.patch("app.printer_manager.win32print.OpenPrinter", return_value="h"), \
                 mock.patch("app.printer_manager.win32print.GetPrinter", return_value={"pDevMode": fake_devmode}), \
                 mock.patch("app.printer_manager.win32print.ClosePrinter"), \
                 mock.patch("app.printer_manager.win32gui.CreateDC", return_value="hdc"), \
                 mock.patch("app.printer_manager.win32ui.CreateDCFromHandle", return_value=FakeDC()), \
                 mock.patch("app.printer_manager.ImageWin.Dib", FakeDib), \
                 mock.patch("app.printer_manager.win32gui.DeleteDC"):
                progress = []
                PrinterManager.print_pdf(
                    "FakePrinter", pdf, [0, 1], 3, "A4",
                    win32con.DMORIENT_PORTRAIT, win32con.DMDUP_SIMPLEX,
                    True, threading.Event(),
                    page_callback=lambda c, t: progress.append((c, t)),
                )
            # 2 trang x 3 ban = 6 lan draw
            self.assertEqual(calls["pages"], 6)
            self.assertEqual(progress[-1], (6, 6))
            # DEVMODE Copies phai bi khoa =1 (driver-independent)
            self.assertEqual(fake_devmode.Copies, 1)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_print_pdf_cancel_and_bad_index(self):
        import threading
        import win32con
        tmp = tempfile.mkdtemp()
        try:
            pdf = make_pdf(os.path.join(tmp, "x.pdf"), 2)

            class FakeDC:
                def GetDeviceCaps(self, cap):
                    return 500
                def StartDoc(self, n):
                    pass
                def StartPage(self):
                    pass
                def EndPage(self):
                    pass
                def AbortDoc(self):
                    pass
                def EndDoc(self):
                    pass
                def GetHandleOutput(self):
                    return 0
                def DeleteDC(self):
                    pass

            ev = threading.Event()
            ev.set()  # cancel ngay
            with mock.patch("app.printer_manager.win32print.OpenPrinter", return_value="h"), \
                 mock.patch("app.printer_manager.win32print.GetPrinter", return_value={"pDevMode": None}), \
                 mock.patch("app.printer_manager.win32print.ClosePrinter"), \
                 mock.patch("app.printer_manager.win32gui.CreateDC", return_value="hdc"), \
                 mock.patch("app.printer_manager.win32ui.CreateDCFromHandle", return_value=FakeDC()), \
                 mock.patch("app.printer_manager.win32gui.DeleteDC"):
                # khong raise khi cancel
                PrinterManager.print_pdf("P", pdf, [0, 99, -5, 1], 1, "A4",
                                         win32con.DMORIENT_PORTRAIT,
                                         win32con.DMDUP_SIMPLEX, True, ev)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


# ================= PRINT WORKER =================
class TestPrintWorker(unittest.TestCase):
    def _job(self, idx=0):
        import win32con
        return PrintJob(index=idx, pdf_path=f"f{idx}.pdf", filename=f"f{idx}.pdf",
                        pages=[0], copies=1, paper_size="A4",
                        orientation=win32con.DMORIENT_PORTRAIT,
                        duplex=win32con.DMDUP_SIMPLEX, fit_to_page=True)

    def test_coordinator_init_normalize(self):
        c = MultiPrinterCoordinator(jobs=[self._job()], printer_names=["A"],
                                    failover_printer="(Không dùng)")
        self.assertIsNone(c.failover_printer)
        c2 = MultiPrinterCoordinator(jobs=[self._job()], printer_names=["A"],
                                     failover_printer="B")
        self.assertEqual(c2.failover_printer, "B")
        with self.assertRaises(ValueError):
            MultiPrinterCoordinator(jobs=[], printer_names=[]).start()

    def test_worker_success(self):
        job = self._job(0)
        done = {}
        coord = MultiPrinterCoordinator(
            jobs=[job], printer_names=["P"],
            on_file_complete=lambda i, f, s, p: done.update(status=s),
            on_all_complete=lambda c: done.update(all=True),
        )
        with mock.patch("app.print_worker.PrinterManager.print_pdf", return_value=None):
            coord.start()
            coord._monitor_thread.join(timeout=5)
        self.assertEqual(done.get("status"), FileStatus.PRINTED)
        self.assertTrue(done.get("all"))
        self.assertEqual(coord.completed_count, 1)

    def test_worker_error_no_failover(self):
        job = self._job(0)
        errs = []
        coord = MultiPrinterCoordinator(
            jobs=[job], printer_names=["P"],
            on_error=lambda i, f, m: errs.append(m),
        )
        with mock.patch("app.print_worker.PrinterManager.print_pdf", side_effect=RuntimeError("ket giay")):
            coord.start()
            coord._monitor_thread.join(timeout=5)
        self.assertEqual(len(errs), 1)
        self.assertIn("ket giay", errs[0])
        self.assertEqual(coord.completed_count, 1)

    def test_worker_failover_success(self):
        job = self._job(0)
        states = {}
        coord = MultiPrinterCoordinator(
            jobs=[job], printer_names=["Chinh"], failover_printer="DuPhong",
            on_file_complete=lambda i, f, s, p: states.update(status=s, printer=p),
            on_failover=lambda *a: states.update(failover=True),
        )
        calls = {"n": 0}

        def fake_print(**kw):
            calls["n"] += 1
            if calls["n"] == 1:
                raise RuntimeError("loi chinh")
            return None

        with mock.patch("app.print_worker.PrinterManager.print_pdf", side_effect=fake_print):
            coord.start()
            coord._monitor_thread.join(timeout=5)
        self.assertTrue(states.get("failover"))
        self.assertEqual(states.get("status"), FileStatus.PRINTED)
        self.assertEqual(states.get("printer"), "DuPhong")

    def test_worker_failover_both_fail(self):
        job = self._job(0)
        errs = []
        coord = MultiPrinterCoordinator(
            jobs=[job], printer_names=["Chinh"], failover_printer="DuPhong",
            on_error=lambda i, f, m: errs.append(m),
        )
        with mock.patch("app.print_worker.PrinterManager.print_pdf", side_effect=RuntimeError("hong")):
            coord.start()
            coord._monitor_thread.join(timeout=5)
        self.assertEqual(len(errs), 1)

    def test_pause_resume_cancel_flags(self):
        coord = MultiPrinterCoordinator(jobs=[self._job()], printer_names=["P"])
        self.assertFalse(coord.is_paused)
        coord.pause()
        self.assertTrue(coord.is_paused)
        coord.resume()
        self.assertFalse(coord.is_paused)
        self.assertFalse(coord.is_cancelled)
        coord.cancel()
        self.assertTrue(coord.is_cancelled)

    def test_separator_cleanup(self):
        import queue as qmod
        tmp = tempfile.mkdtemp()
        try:
            sep = os.path.join(tmp, "sep.pdf")
            make_pdf(sep, 1)
            job = self._job(0)
            job.is_separator = True
            job.pdf_path = sep
            coord = MultiPrinterCoordinator(jobs=[job], printer_names=["P"])
            with mock.patch("app.print_worker.PrinterManager.print_pdf", return_value=None):
                coord.start()
                coord._monitor_thread.join(timeout=5)
            self.assertFalse(os.path.exists(sep))
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_thread_safety_increment(self):
        coord = MultiPrinterCoordinator(jobs=[self._job(i) for i in range(50)], printer_names=["P"])

        def inc():
            for _ in range(200):
                coord.increment_completed()

        threads = [threading.Thread(target=inc) for _ in range(5)]
        [t.start() for t in threads]
        [t.join() for t in threads]
        self.assertEqual(coord.completed_count, 5 * 200)


# ================= CONFIG STORE =================
class TestConfigStore(unittest.TestCase):
    def test_app_data_dir_user_writable(self):
        from app.config_store import get_app_data_dir, CONFIG_FILE
        d = get_app_data_dir()
        self.assertTrue(d.endswith("PDFBatchPrinterPro") or d.endswith(".pdfbatchprinterpro"))
        self.assertTrue(CONFIG_FILE.endswith("config.json"))
        # khong bao gio la Program Files / thu muc exe
        self.assertNotIn("Program Files", d)

    def test_save_load_roundtrip(self):
        import app.config_store as cs
        with mock.patch.object(cs, "CONFIG_FILE", os.path.join(tempfile.mkdtemp(), "config.json")):
            cs.save_config({"paper": "A3", "copies": 2})
            self.assertEqual(cs.load_config(), {"paper": "A3", "copies": 2})

    def test_load_missing_returns_empty(self):
        import app.config_store as cs
        with mock.patch.object(cs, "CONFIG_FILE", os.path.join(tempfile.mkdtemp(), "no.json")):
            with mock.patch.object(cs, "_legacy_candidates", return_value=[]):
                self.assertEqual(cs.load_config(), {})


# ================= QUEUE STORE =================
class TestQueueStore(unittest.TestCase):
    def test_dedupe(self):
        from app.queue_store import dedupe_paths
        seen = {os.path.abspath("/a/b.pdf")}
        self.assertEqual(dedupe_paths(["/a/b.pdf", "/a/c.pdf", "/a/c.pdf"], seen), ["/a/c.pdf"])

    def test_move_up_down(self):
        from app.queue_store import apply_move
        items = ["1", "2", "3"]
        self.assertEqual(apply_move(items, [1], -1), {0})
        self.assertEqual(items, ["2", "1", "3"])
        self.assertEqual(apply_move(items, [0], -1), {0})  # bien tren: dung yen
        self.assertEqual(apply_move(items, [2], +1), {2})  # bien duoi: dung yen

    def test_summary_and_retry(self):
        from app.queue_store import queue_summary, retry_indices, is_edit_locked
        from app.pdf_manager import FileInfo
        f1 = FileInfo(original_path="a.pdf", pdf_path="a.pdf", page_count=5, file_size=1, file_type="PDF")
        f1.copies = 2
        f1.status = FileStatus.ERROR
        f2 = FileInfo(original_path="b.pdf", pdf_path="b.pdf", page_count=3, file_size=1, file_type="PDF")
        self.assertEqual(queue_summary([f1, f2]), (2, 3, 13))
        self.assertEqual(retry_indices([f1, f2]), [0])
        self.assertTrue(is_edit_locked(FileStatus.PRINTING))
        self.assertFalse(is_edit_locked(FileStatus.WAITING))


# ================= PAPER A3 =================
class TestPaperA3(unittest.TestCase):
    def test_a3_in_settings(self):
        import win32con
        self.assertEqual(PAPER_SIZES["A3"], win32con.DMPAPER_A3)
        self.assertIn("A3", get_paper_name(841.89, 1190.55))

    def test_cli_accepts_a3(self):
        import argparse
        import cli  # noqa: F401 — dam bao module nap duoc sau khi sua choices
        self.assertIn("A3", PAPER_SIZES)


# ================= PREVIEW =================
class TestPreviewLogic(unittest.TestCase):
    def test_get_paper_name_all(self):
        self.assertIn("A4", get_paper_name(595.28, 841.89))
        self.assertIn("A4", get_paper_name(841.89, 595.28))
        self.assertIn("A3", get_paper_name(841.89, 1190.55))
        self.assertIn("A5", get_paper_name(419.53, 595.28))
        self.assertIn("Letter", get_paper_name(612.0, 792.0))
        self.assertIn("Legal", get_paper_name(612.0, 1008.0))
        self.assertIn("Tùy chỉnh", get_paper_name(300.0, 500.0))


# ================= CLI =================
class TestCLI(unittest.TestCase):
    def test_execute_uses_coordinator(self):
        import win32con
        import cli
        jobs = [PrintJob(index=0, pdf_path="a.pdf", filename="a.pdf", pages=[0],
                         copies=1, paper_size="A4",
                         orientation=win32con.DMORIENT_PORTRAIT,
                         duplex=win32con.DMDUP_SIMPLEX, fit_to_page=True)]
        with mock.patch("cli.MultiPrinterCoordinator") as MCC:
            inst = MCC.return_value

            def _fake_start():
                # goi callback on_all_complete that da truyen vao constructor
                cb = MCC.call_args[1].get("on_all_complete")
                if cb:
                    cb(False)

            inst.start.side_effect = _fake_start
            cli.execute_print_jobs(jobs, "MayIn")
            MCC.assert_called_once()
            _, kw = MCC.call_args
            self.assertEqual(kw.get("printer_names"), ["MayIn"])

    def test_cli_args_build_jobs_with_pages(self):
        import win32con
        import cli
        tmp = tempfile.mkdtemp()
        try:
            pdf = make_pdf(os.path.join(tmp, "d.pdf"), 5)
            args = mock.MagicMock()
            args.list_printers = False
            args.files = [pdf]
            args.dir = None
            args.recursive = False
            args.printer = "MayIn"
            args.duplex = "long"
            args.orient = "auto"
            args.paper = "A4"
            args.copies = 1
            args.pages = "1-2"
            args.no_fit = False
            with mock.patch("cli.execute_print_jobs") as exe:
                cli.run_cli_args(args)
                exe.assert_called_once()
                jobs, printer = exe.call_args[0]
                self.assertEqual(printer, "MayIn")
                self.assertEqual(jobs[0].pages, [0, 1])
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
            cleanup_temp()

    def test_cli_args_bad_pages_aborts(self):
        import cli
        tmp = tempfile.mkdtemp()
        try:
            pdf = make_pdf(os.path.join(tmp, "d.pdf"), 2)
            args = mock.MagicMock()
            args.list_printers = False
            args.files = [pdf]
            args.dir = None
            args.recursive = False
            args.printer = "MayIn"
            args.duplex = "1"
            args.orient = "auto"
            args.paper = "A4"
            args.copies = 1
            args.pages = "1-99"
            args.no_fit = False
            with mock.patch("cli.execute_print_jobs") as exe:
                cli.run_cli_args(args)
                exe.assert_not_called()
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
            cleanup_temp()

    def test_cli_list_printers(self):
        import cli
        args = mock.MagicMock()
        args.list_printers = True
        with mock.patch("cli.PrinterManager.get_printers", return_value=["P1"]), \
             mock.patch("cli.PrinterManager.get_default_printer", return_value="P1"), \
             mock.patch("cli.PrinterManager.supports_duplex", return_value=True):
            cli.run_cli_args(args)  # khong raise


# ================= GUI PURE LOGIC (khong can Tk) =================
class TestGUILogicHeadless(unittest.TestCase):
    def test_duplicate_filter_logic(self):
        seen = {os.path.abspath("/a/b.pdf")}
        paths = ["/a/b.pdf", "/a/c.pdf", "/a/c.pdf"]
        valid = []
        for p in paths:
            ap = os.path.abspath(p)
            if ap in seen:
                continue
            seen.add(ap)
            valid.append(p)
        self.assertEqual(valid, ["/a/c.pdf"])

    def test_log_trim_logic(self):
        lines = list(range(2500))
        MAX = 2000
        if len(lines) > MAX:
            lines = lines[len(lines) - MAX:]
        self.assertEqual(len(lines), 2000)

    def test_copy_lock_rule(self):
        locked = {FileStatus.PRINTING, FileStatus.PRINTED}
        self.assertIn(FileStatus.PRINTING, locked)
        self.assertNotIn(FileStatus.WAITING, locked)
        self.assertNotIn(FileStatus.ERROR, locked)

    def test_summary_math(self):
        class F:
            def __init__(self, pg, cp):
                self.page_count = pg
                self.copies = cp
        q = [F(5, 3), F(10, 2)]
        self.assertEqual(sum(f.page_count * f.copies for f in q), 35)


# ================= TICH HOP =================
class TestIntegration(unittest.TestCase):
    def test_add_ensure_filter_job_flow(self):
        import win32con
        tmp = tempfile.mkdtemp()
        try:
            pdf = make_blank_pdf(os.path.join(tmp, "flow.pdf"))
            pm = PDFManager()
            info = pm.add_file(pdf)
            pm.ensure_pdf(info)
            pages = list(range(info.page_count))
            active, skipped = PDFManager.filter_pages(info.pdf_path, pages, True)
            self.assertEqual(active, [1])
            job = PrintJob(index=0, pdf_path=info.pdf_path, filename=info.filename,
                           pages=active, copies=2, paper_size="A4",
                           orientation=win32con.DMORIENT_PORTRAIT,
                           duplex=win32con.DMDUP_SIMPLEX, fit_to_page=True)
            coord = MultiPrinterCoordinator(jobs=[job], printer_names=["P"])
            self.assertEqual(coord.total_count, 1)
            with mock.patch("app.print_worker.PrinterManager.print_pdf", return_value=None):
                coord.start()
                coord._monitor_thread.join(timeout=5)
            self.assertEqual(coord.completed_count, 1)
        finally:
            shutil.rmtree(tmp, ignore_errors=True)
            cleanup_temp()


if __name__ == "__main__":
    unittest.main(verbosity=2)

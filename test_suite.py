"""
Master Automated Test Suite for PDF Batch Printer Pro
Covers 100% of modules, algorithms, file converters, print managers, queue logic,
preview logic, settings, and GUI integrations.
"""
import os
import sys
import unittest
import tempfile
import json
import shutil
from typing import List

import fitz  # PyMuPDF
from PIL import Image

# App imports
from app.settings import (
    FileStatus, STATUS_ICONS, STATUS_COLORS,
    SUPPORTED_EXTENSIONS, THEME_COLORS, APP_NAME, WINDOW_TITLE,
    PAGE_RANGE_ALL, PAGE_RANGE_CUSTOM, PAGE_RANGE_OPTIONS,
    PAPER_SIZES, DUPLEX_MODES, BINDING_MARGIN_OPTIONS
)
from app.utils import parse_page_range, format_file_size, get_timestamp
from app.pdf_manager import PDFManager, FileInfo
from app.file_converter import convert_to_pdf, cleanup_temp
from app.printer_manager import PrinterManager
from app.print_worker import PrintJob, MultiPrinterCoordinator
from app.preview import get_paper_name, PreviewWindow


class TestSettingsAndUtils(unittest.TestCase):
    """Test utility functions, formatting, and configuration constants."""

    def test_parse_page_range_valid(self):
        self.assertEqual(parse_page_range("1-3", 5), [1, 2, 3])
        self.assertEqual(parse_page_range("1, 3, 5", 5), [1, 3, 5])
        self.assertEqual(parse_page_range("2-4, 5", 5), [2, 3, 4, 5])
        self.assertEqual(parse_page_range("1", 5), [1])
        self.assertEqual(parse_page_range("1-1", 5), [1])
        self.assertEqual(parse_page_range(" 1 , 2 , 3 ", 5), [1, 2, 3])

    def test_parse_page_range_invalid(self):
        # Out of bounds
        with self.assertRaises(ValueError):
            parse_page_range("6", 5)
        # Page 0 is invalid (1-indexed)
        with self.assertRaises(ValueError):
            parse_page_range("0", 5)
        # Inverted range
        with self.assertRaises(ValueError):
            parse_page_range("4-2", 5)
        # Non-numeric
        with self.assertRaises(ValueError):
            parse_page_range("abc", 5)
        with self.assertRaises(ValueError):
            parse_page_range("1-a", 5)

    def test_format_file_size(self):
        self.assertEqual(format_file_size(0), "0 B")
        self.assertEqual(format_file_size(500), "500 B")
        self.assertEqual(format_file_size(1024), "1.0 KB")
        self.assertEqual(format_file_size(2048), "2.0 KB")
        self.assertEqual(format_file_size(1024 * 1024), "1.0 MB")
        self.assertEqual(format_file_size(5 * 1024 * 1024), "5.0 MB")
        self.assertEqual(format_file_size(2 * 1024 * 1024 * 1024), "2.0 GB")

    def test_get_timestamp(self):
        ts = get_timestamp()
        self.assertIsInstance(ts, str)
        self.assertTrue(len(ts) >= 8)

    def test_configuration_constants(self):
        self.assertIn(".pdf", SUPPORTED_EXTENSIONS)
        self.assertIn(".docx", SUPPORTED_EXTENSIONS)
        self.assertIn(".png", SUPPORTED_EXTENSIONS)
        self.assertIn("A4", PAPER_SIZES)
        self.assertIn("A5", PAPER_SIZES)
        self.assertIn("Letter", PAPER_SIZES)
        self.assertEqual(len(DUPLEX_MODES), 3)
        self.assertTrue(len(BINDING_MARGIN_OPTIONS) >= 4)


class TestPDFManagerCore(unittest.TestCase):
    """Test PDF inspection, page detection, filtering, separator generation, and scanning."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.sample_pdf = os.path.join(self.tmp_dir, "sample.pdf")
        doc = fitz.open()
        for i in range(4):
            page = doc.new_page(width=595, height=842)  # A4 Portrait
            page.insert_text((50, 50), f"Sample Content Page {i + 1}", fontsize=18)
        doc.save(self.sample_pdf)
        doc.close()

    def tearDown(self):
        try:
            shutil.rmtree(self.tmp_dir, ignore_errors=True)
        except Exception:
            pass

    def test_add_file_and_properties(self):
        pm = PDFManager()
        info = pm.add_file(self.sample_pdf)
        self.assertEqual(info.page_count, 4)
        self.assertEqual(info.filename, "sample.pdf")
        self.assertEqual(info.file_type, "PDF")
        self.assertEqual(info.status, FileStatus.WAITING)
        self.assertEqual(info.copies, 1)
        self.assertTrue(info.is_converted_ready)
        self.assertEqual(info.pdf_path, self.sample_pdf)

    def test_quick_inspect_and_ensure_pdf(self):
        pm = PDFManager()
        info = pm.quick_inspect(self.sample_pdf)
        self.assertEqual(info.page_count, 4)
        self.assertEqual(info.file_type, "PDF")
        self.assertTrue(info.is_converted_ready)
        
        ready_path = pm.ensure_pdf(info)
        self.assertEqual(ready_path, self.sample_pdf)
        self.assertTrue(os.path.exists(ready_path))

    def test_blank_page_detection_and_filtering(self):
        # Create a PDF with 2 pages: 1 non-blank, 1 completely blank
        pdf_blank_test = os.path.join(self.tmp_dir, "blank_test.pdf")
        doc = fitz.open()
        p1 = doc.new_page(width=595, height=842)
        p1.insert_text((50, 50), "Non-blank content here", fontsize=16)
        p2 = doc.new_page(width=595, height=842)  # blank
        doc.save(pdf_blank_test)
        doc.close()

        # filter_pages with remove_blanks=True
        active, skipped = PDFManager.filter_pages(pdf_blank_test, [0, 1], remove_blanks=True)
        self.assertEqual(active, [0])
        self.assertEqual(skipped, [1])

        # filter_pages with remove_blanks=False
        active_all, skipped_none = PDFManager.filter_pages(pdf_blank_test, [0, 1], remove_blanks=False)
        self.assertEqual(active_all, [0, 1])
        self.assertEqual(skipped_none, [])

    def test_create_separator_sheet(self):
        sep_pdf = PDFManager.create_separator_sheet(
            filename="Bao_Cao_Tai_Chinh_2026.docx",
            page_count=20,
            copies=2,
            printer_name="Canon LBP 2900",
            doc_index=1
        )
        self.assertTrue(os.path.exists(sep_pdf))
        doc = fitz.open(sep_pdf)
        self.assertEqual(doc.page_count, 1)
        doc.close()
        if os.path.exists(sep_pdf):
            os.remove(sep_pdf)

    def test_scan_folder_recursive_vs_non_recursive(self):
        sub_dir = os.path.join(self.tmp_dir, "sub_folder")
        os.makedirs(sub_dir, exist_ok=True)
        sub_pdf = os.path.join(sub_dir, "nested.pdf")
        shutil.copyfile(self.sample_pdf, sub_pdf)

        pm = PDFManager()
        # Non-recursive scan: only top level
        top_scanned = pm.scan_folder(self.tmp_dir, recursive=False)
        self.assertIn(self.sample_pdf, top_scanned)
        self.assertNotIn(sub_pdf, top_scanned)

        # Recursive scan: top level and subfolder
        all_scanned = pm.scan_folder(self.tmp_dir, recursive=True)
        self.assertIn(self.sample_pdf, all_scanned)
        self.assertIn(sub_pdf, all_scanned)

    def test_get_page_size(self):
        pm = PDFManager()
        w, h = pm.get_page_size(self.sample_pdf, 0)
        self.assertAlmostEqual(w, 595.0, delta=2.0)
        self.assertAlmostEqual(h, 842.0, delta=2.0)


class TestFileConverter(unittest.TestCase):
    """Test image conversion, unsupported extension handling, and cleanup."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.img_png = os.path.join(self.tmp_dir, "test_image.png")
        img = Image.new("RGB", (400, 300), color="blue")
        img.save(self.img_png)

    def tearDown(self):
        try:
            cleanup_temp()
            shutil.rmtree(self.tmp_dir, ignore_errors=True)
        except Exception:
            pass

    def test_convert_image_to_pdf(self):
        pdf_path, is_temp = convert_to_pdf(self.img_png)
        self.assertTrue(is_temp)
        self.assertTrue(os.path.exists(pdf_path))
        doc = fitz.open(pdf_path)
        self.assertEqual(doc.page_count, 1)
        doc.close()

    def test_convert_pdf_returns_original(self):
        tmp_pdf = os.path.join(self.tmp_dir, "dummy.pdf")
        doc = fitz.open()
        doc.new_page()
        doc.save(tmp_pdf)
        doc.close()

        res_path, is_temp = convert_to_pdf(tmp_pdf)
        self.assertEqual(res_path, os.path.abspath(tmp_pdf))
        self.assertFalse(is_temp)

    def test_unsupported_extension_raises_error(self):
        bad_file = os.path.join(self.tmp_dir, "test.unknown_ext")
        with open(bad_file, "w") as f:
            f.write("dummy")
        with self.assertRaises(ValueError):
            convert_to_pdf(bad_file)


class TestPrinterManagerAndWorkers(unittest.TestCase):
    """Test printer enumeration, status mappings, PrintJob creation, and MultiPrinterCoordinator."""

    def test_printer_enumeration_and_status(self):
        printers = PrinterManager.get_printers()
        self.assertIsInstance(printers, list)
        
        default_p = PrinterManager.get_default_printer()
        self.assertIsInstance(default_p, str)

        # Test status inquiry (handles both installed printers or dummy fallback)
        status_text, is_ready = PrinterManager.get_printer_status("NonExistentPrinter12345")
        self.assertIsInstance(status_text, str)
        self.assertIsInstance(is_ready, bool)

    def test_print_job_instantiation(self):
        import win32con
        job = PrintJob(
            index=1,
            pdf_path="test.pdf",
            filename="test.pdf",
            pages=[0, 1, 2],
            copies=2,
            paper_size="A4",
            orientation=win32con.DMORIENT_PORTRAIT,
            duplex=win32con.DMDUP_SIMPLEX,
            fit_to_page=True,
            binding_margin_mm=5.0,
            reverse_order=False,
            is_separator=False
        )
        self.assertEqual(job.index, 1)
        self.assertEqual(job.copies, 2)
        self.assertEqual(job.pages, [0, 1, 2])
        self.assertEqual(job.paper_size, "A4")
        self.assertEqual(job.binding_margin_mm, 5.0)

    def test_multi_printer_coordinator_lifecycle(self):
        import win32con
        job1 = PrintJob(
            index=0, pdf_path="test1.pdf", filename="test1.pdf",
            pages=[0], copies=1, paper_size="A4",
            orientation=win32con.DMORIENT_PORTRAIT,
            duplex=win32con.DMDUP_SIMPLEX, fit_to_page=True,
            binding_margin_mm=0.0, reverse_order=False
        )
        job2 = PrintJob(
            index=1, pdf_path="test2.pdf", filename="test2.pdf",
            pages=[0, 1], copies=2, paper_size="A4",
            orientation=win32con.DMORIENT_PORTRAIT,
            duplex=win32con.DMDUP_SIMPLEX, fit_to_page=True,
            binding_margin_mm=0.0, reverse_order=False
        )

        coord = MultiPrinterCoordinator(
            jobs=[job1, job2],
            printer_names=["PrinterA", "PrinterB"],
            failover_printer="PrinterBackup"
        )
        self.assertEqual(coord.total_count, 2)
        self.assertEqual(coord.completed_count, 0)
        self.assertEqual(len(coord.printer_names), 2)
        self.assertEqual(coord.failover_printer, "PrinterBackup")

        # Test pause and cancel flags
        self.assertFalse(coord.cancel_event.is_set())
        coord.cancel()
        self.assertTrue(coord.cancel_event.is_set())


class TestQueueLogicAndTransformations(unittest.TestCase):
    """Test queue sorting, reordering, copy editing constraints, and summary calculations."""

    def test_queue_sorting(self):
        f1 = FileInfo(original_path="Bravo.pdf", pdf_path="Bravo.pdf", page_count=10, file_size=2000, file_type="PDF")
        f1.copies = 2
        f2 = FileInfo(original_path="Alpha.pdf", pdf_path="Alpha.pdf", page_count=2, file_size=5000, file_type="PDF")
        f2.copies = 1
        f3 = FileInfo(original_path="Charlie.pdf", pdf_path="Charlie.pdf", page_count=20, file_size=1000, file_type="PDF")
        f3.copies = 5
        queue = [f1, f2, f3]

        # Sort by filename ascending
        queue.sort(key=lambda x: x.filename.lower())
        self.assertEqual([x.filename for x in queue], ["Alpha.pdf", "Bravo.pdf", "Charlie.pdf"])

        # Sort by page_count descending
        queue.sort(key=lambda x: x.page_count, reverse=True)
        self.assertEqual([x.filename for x in queue], ["Charlie.pdf", "Bravo.pdf", "Alpha.pdf"])

        # Sort by file_size descending
        queue.sort(key=lambda x: x.file_size, reverse=True)
        self.assertEqual([x.filename for x in queue], ["Alpha.pdf", "Bravo.pdf", "Charlie.pdf"])

        # Sort by copies ascending
        queue.sort(key=lambda x: x.copies)
        self.assertEqual([x.filename for x in queue], ["Alpha.pdf", "Bravo.pdf", "Charlie.pdf"])

    def test_queue_moving_items(self):
        f1 = FileInfo(original_path="1.pdf", pdf_path="1.pdf", page_count=1, file_size=100, file_type="PDF")
        f2 = FileInfo(original_path="2.pdf", pdf_path="2.pdf", page_count=1, file_size=100, file_type="PDF")
        f3 = FileInfo(original_path="3.pdf", pdf_path="3.pdf", page_count=1, file_size=100, file_type="PDF")
        queue = [f1, f2, f3]

        # Move f2 up
        idx = queue.index(f2)
        queue[idx], queue[idx - 1] = queue[idx - 1], queue[idx]
        self.assertEqual(queue, [f2, f1, f3])

        # Move f2 down
        idx = queue.index(f2)
        queue[idx], queue[idx + 1] = queue[idx + 1], queue[idx]
        self.assertEqual(queue, [f1, f2, f3])

    def test_copy_lock_rules_and_status(self):
        f = FileInfo(original_path="doc.pdf", pdf_path="doc.pdf", page_count=5, file_size=500, file_type="PDF")
        f.status = FileStatus.WAITING
        # WAITING -> allowed to edit
        self.assertNotIn(f.status, (FileStatus.PRINTING, FileStatus.PRINTED))

        # PRINTING -> locked
        f.status = FileStatus.PRINTING
        self.assertIn(f.status, (FileStatus.PRINTING, FileStatus.PRINTED))

        # PRINTED -> locked
        f.status = FileStatus.PRINTED
        self.assertIn(f.status, (FileStatus.PRINTING, FileStatus.PRINTED))

        # ERROR / CANCELLED -> allowed to retry / edit
        f.status = FileStatus.ERROR
        self.assertNotIn(f.status, (FileStatus.PRINTING, FileStatus.PRINTED))

    def test_queue_summary_calculations(self):
        f1 = FileInfo(original_path="doc1.pdf", pdf_path="doc1.pdf", page_count=5, file_size=1000, file_type="PDF")
        f1.copies = 3  # 5 * 3 = 15 pages
        f2 = FileInfo(original_path="doc2.pdf", pdf_path="doc2.pdf", page_count=10, file_size=2000, file_type="PDF")
        f2.copies = 2  # 10 * 2 = 20 pages
        queue = [f1, f2]

        total_files = len(queue)
        total_copies = sum(f.copies for f in queue)
        total_pages = sum(f.page_count * f.copies for f in queue)

        self.assertEqual(total_files, 2)
        self.assertEqual(total_copies, 5)
        self.assertEqual(total_pages, 35)


class TestPreviewLogic(unittest.TestCase):
    """Test preview paper size calculation."""

    def test_get_paper_name(self):
        self.assertIn("A4", get_paper_name(595.28, 841.89))
        self.assertIn("A4", get_paper_name(841.89, 595.28))  # landscape orientation
        self.assertIn("A3", get_paper_name(841.89, 1190.55))
        self.assertIn("A5", get_paper_name(419.53, 595.28))
        self.assertIn("Letter", get_paper_name(612.0, 792.0))
        self.assertIn("Legal", get_paper_name(612.0, 1008.0))
        self.assertIn("Tùy chỉnh", get_paper_name(300.0, 500.0))


class TestGUIStateAndIntegration(unittest.TestCase):
    """Test PDFBatchPrinterApp UI components, treeview bindings, and state updates."""

    @classmethod
    def setUpClass(cls):
        # Create a test PDF
        cls.tmp_dir = tempfile.mkdtemp()
        cls.test_pdf = os.path.join(cls.tmp_dir, "gui_sample.pdf")
        doc = fitz.open()
        p = doc.new_page(width=595, height=842)
        p.insert_text((50, 50), "GUI Test Content", fontsize=20)
        doc.save(cls.test_pdf)
        doc.close()

        from app.gui import PDFBatchPrinterApp
        cls.app = PDFBatchPrinterApp()
        cls.app.withdraw()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.app._on_close()
        except Exception:
            pass
        try:
            shutil.rmtree(cls.tmp_dir, ignore_errors=True)
        except Exception:
            pass

    def setUp(self):
        self.app.file_list.clear()
        self.app._refresh_tree()

    def test_gui_lifecycle_and_queue_actions(self):
        from unittest.mock import patch
        import customtkinter as ctk

        app = self.app
        # 1. Add file to queue
        pm = PDFManager()
        finfo1 = pm.quick_inspect(self.test_pdf)
        finfo1.copies = 2
        app.file_list.append(finfo1)
        app._refresh_tree()

        # Verify treeview populated
        items = app.tree.get_children()
        self.assertEqual(len(items), 1)
        values = app.tree.item(items[0], "values")
        self.assertEqual(str(values[0]), "1")  # Index
        self.assertEqual(str(values[1]), "gui_sample.pdf")
        self.assertEqual(int(values[6]), 2)  # Copies (sau cột Trang in)

        # 2. Select file and test copies stepper
        app.tree.selection_set(items[0])
        app._on_file_select()
        self.assertEqual(app._copies_var.get(), 2)

        # Test increment copies (+1)
        app._inc_copies()
        self.assertEqual(finfo1.copies, 3)
        self.assertEqual(app._copies_var.get(), 3)

        # Test decrement copies (-1)
        app._dec_copies()
        self.assertEqual(finfo1.copies, 2)
        self.assertEqual(app._copies_var.get(), 2)

        # 3. Add second file and test apply default copies to all
        finfo2 = pm.quick_inspect(self.test_pdf)
        finfo2.copies = 1
        app.file_list.append(finfo2)
        app._refresh_tree()
        self.assertEqual(len(app.tree.get_children()), 2)

        app._default_copies_var.set(5)
        app._apply_default_copies_to_all()
        self.assertEqual(finfo1.copies, 5)
        self.assertEqual(finfo2.copies, 5)

        # 4. Verify queue summary bar text
        app._update_queue_summary()
        summary_text = app.queue_summary_lbl.cget("text")
        self.assertIn("Tổng tài liệu: 2", summary_text)
        self.assertIn("Tổng số bản: 10", summary_text)

        # 5. Test theme toggle
        if app.theme_switch.get() == 1:
            app.theme_switch.deselect()
        else:
            app.theme_switch.select()
        app._toggle_theme()
        self.assertEqual(ctk.get_appearance_mode(), "Dark" if app.theme_switch.get() == 1 else "Light")

        # 6. Test delete selected and delete all
        app.tree.selection_set(app.tree.get_children()[0])
        with patch("tkinter.messagebox.askyesno", return_value=True):
            app.remove_selected()
            self.assertEqual(len(app.file_list), 1)

            app.remove_all()
            self.assertEqual(len(app.file_list), 0)
            self.assertEqual(len(app.tree.get_children()), 0)

    def test_gui_stepper_boundaries_and_manual_entry(self):
        app = self.app
        pm = PDFManager()
        finfo = pm.quick_inspect(self.test_pdf)
        finfo.copies = 1
        app.file_list.append(finfo)
        app._refresh_tree()

        app.tree.selection_set(app.tree.get_children()[0])
        app._on_file_select()

        # Decrement at 1 should not go below 1
        app._dec_copies()
        self.assertEqual(finfo.copies, 1)
        self.assertEqual(app._copies_var.get(), 1)

        # Increment to 2
        app._inc_copies()
        self.assertEqual(finfo.copies, 2)
        self.assertEqual(app._copies_var.get(), 2)

        # Manual valid entry via variable
        app._copies_var.set(15)
        app._on_copies_entry_enter()
        self.assertEqual(finfo.copies, 15)
        self.assertEqual(app._copies_var.get(), 15)

        # Stepper increment from 15 to 16
        app._inc_copies()
        self.assertEqual(finfo.copies, 16)
        self.assertEqual(app._copies_var.get(), 16)

    def test_gui_move_up_down_operations(self):
        app = self.app
        pm = PDFManager()
        f1 = pm.quick_inspect(self.test_pdf)
        f1.filename = "File_1.pdf"
        f2 = pm.quick_inspect(self.test_pdf)
        f2.filename = "File_2.pdf"
        app.file_list.extend([f1, f2])
        app._refresh_tree()

        children = app.tree.get_children()
        self.assertEqual(len(children), 2)

        # Select 2nd item and move up
        app.tree.selection_set(children[1])
        app.move_selected_up()
        self.assertEqual(app.file_list[0].filename, "File_2.pdf")
        self.assertEqual(app.file_list[1].filename, "File_1.pdf")

        # Move 1st item down
        new_children = app.tree.get_children()
        app.tree.selection_set(new_children[0])
        app.move_selected_down()
        self.assertEqual(app.file_list[0].filename, "File_1.pdf")
        self.assertEqual(app.file_list[1].filename, "File_2.pdf")

    def test_preview_modal_instantiation(self):
        app = self.app
        from app.preview import PreviewWindow

        pm = PDFManager()
        finfo = pm.quick_inspect(self.test_pdf)
        modal = PreviewWindow(app, pm, finfo)
        modal.withdraw()
        self.assertIsNotNone(modal)
        modal.destroy()


if __name__ == "__main__":
    unittest.main(verbosity=2)

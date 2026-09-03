import os
import shutil
import subprocess
import tempfile
import threading
import time
import unittest

import uno

import main


class WriterDocumentTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.profile = tempfile.mkdtemp(prefix="writer-ai-lo-")
        cls.port = 2083
        profile_url = uno.systemPathToFileUrl(cls.profile)
        cls.process = subprocess.Popen([
            "/opt/libreoffice26.2/program/soffice",
            "--headless",
            "--nologo",
            "--nofirststartwizard",
            "--norestore",
            f"-env:UserInstallation={profile_url}",
            f"--accept=socket,host=127.0.0.1,port={cls.port};urp;StarOffice.ServiceManager",
        ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context
        )
        connection = f"uno:socket,host=127.0.0.1,port={cls.port};urp;StarOffice.ComponentContext"
        for _ in range(40):
            try:
                cls.context = resolver.resolve(connection)
                break
            except Exception:
                time.sleep(0.25)
        else:
            cls.process.terminate()
            raise RuntimeError("Unable to connect to headless LibreOffice")

        service_manager = cls.context.ServiceManager
        cls.desktop = service_manager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", cls.context
        )

    @classmethod
    def tearDownClass(cls):
        try:
            cls.desktop.terminate()
        finally:
            cls.process.wait(timeout=10)
            shutil.rmtree(cls.profile, ignore_errors=True)

    def setUp(self):
        hidden = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        hidden.Name = "Hidden"
        hidden.Value = True
        self.doc = self.desktop.loadComponentFromURL(
            "private:factory/swriter", "_blank", 0, (hidden,)
        )
        self.formatter = main.Format(self.context, self.doc)

    def tearDown(self):
        self.doc.close(True)

    def test_document_structure_and_indent(self):
        cursor = self.doc.Text.createTextCursor()
        self.doc.Text.insertString(cursor, "Document title", False)
        self.doc.Text.insertControlCharacter(cursor, uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK"), False)
        self.doc.Text.insertString(cursor, "Body paragraph.", False)

        self.formatter.format_document_structure({
            "infer": True,
            "title": {"bold": True},
            "body": {"first_line_indent": 2},
        })
        paragraphs = list(self.formatter._paragraphs())
        self.assertEqual(paragraphs[0].ParaStyleName, "Title")
        self.assertIn(paragraphs[1].ParaStyleName, {"Text Body", "Standard"})
        self.assertGreater(paragraphs[1].ParaFirstLineIndent, 0)

    def test_numbered_heading_inference(self):
        cursor = self.doc.Text.createTextCursor()
        for index, text in enumerate(("论文标题", "1. 方法", "1.1 数据处理", "这是正文。")):
            if index:
                self.doc.Text.insertControlCharacter(
                    cursor,
                    uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK"),
                    False,
                )
            self.doc.Text.insertString(cursor, text, False)

        self.formatter.format_document_structure({
            "infer": True,
            "title": True,
            "heading_1": True,
            "heading_2": True,
            "body": True,
        })
        paragraphs = list(self.formatter._paragraphs())
        self.assertEqual(paragraphs[1].ParaStyleName, "Heading 1")
        self.assertEqual(paragraphs[2].ParaStyleName, "Heading 2")

    def test_table_formatting(self):
        table = self.doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(2, 2)
        self.doc.Text.insertTextContent(self.doc.Text.createTextCursor(), table, False)
        table.getCellByName("A1").String = "Header"
        table.getCellByName("B1").String = "Value"
        table.getCellByName("A2").String = "Row"
        table.getCellByName("B2").String = "12.5%"

        self.formatter.format_table(table, {
            "header": True,
            "header_background": "0ABAB5",
            "first_column_bold": True,
            "zebra": True,
            "zebra_background": "F2F2F2",
            "align": "center",
            "auto_align": True,
        })
        self.assertEqual(table.getCellByName("A1").BackColor, 0x0ABAB5)
        self.assertEqual(table.getCellByName("A2").BackColor, 0xF2F2F2)
        self.assertEqual(int(table.getCellByName("A1").createTextCursor().ParaAdjust), 3)
        self.formatter.format_table(table, {"auto_align": True})
        self.assertEqual(int(table.getCellByName("B2").createTextCursor().ParaAdjust), 1)

    def test_table_column_widths(self):
        table = self.doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(2, 2)
        self.doc.Text.insertTextContent(self.doc.Text.createTextCursor(), table, False)
        self.formatter.format_table(table, {"column_widths": [30, 70]})
        self.assertAlmostEqual(table.TableColumnSeparators[0].Position, 300, delta=1)

    def test_table_cell_merge(self):
        table = self.doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(2, 2)
        self.doc.Text.insertTextContent(self.doc.Text.createTextCursor(), table, False)
        table.getCellByName("A1").String = "Merged"
        self.formatter.format_table(table, {"merge_cells": [{"start": "A1", "end": "B1"}]})
        self.assertIn("Merged", table.getCellByName("A1").String)
        self.assertNotIn("B1", table.getCellNames())

    def test_table_caption_and_number(self):
        table = self.doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(1, 1)
        self.doc.Text.insertTextContent(self.doc.Text.createTextCursor(), table, False)
        self.formatter.format_table(table, {"caption": "实验结果", "number": 3})
        self.assertIn("表3：实验结果", self.doc.Text.String)

    def test_table_page_break_properties(self):
        table = self.doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(2, 1)
        self.doc.Text.insertTextContent(self.doc.Text.createTextCursor(), table, False)
        self.formatter.format_table(table, {
            "split": False,
            "keep_together": True,
        })
        self.assertFalse(table.Split)
        self.assertTrue(table.KeepTogether)

    def test_empty_document(self):
        paragraphs = list(self.formatter._paragraphs())
        self.assertTrue(all(not paragraph.String for paragraph in paragraphs))

    def test_async_callback_returns_to_writer(self):
        received = []

        class Receiver:
            def _finish_format_request(self, target_doc, data, generation):
                received.append((target_doc, data, generation))

        callback = main.FormatRequestCallback(Receiver(), self.doc, 7)
        async_callback = self.context.ServiceManager.createInstanceWithContext(
            "com.sun.star.awt.AsyncCallback", self.context
        )
        worker = threading.Thread(
            target=lambda: async_callback.addCallback(callback, "callback-ok")
        )
        worker.start()
        worker.join(timeout=2)
        for _ in range(20):
            if received:
                break
            time.sleep(0.1)
        self.assertEqual(received, [(self.doc, "callback-ok", 7)])

    def test_docx_round_trip(self):
        path = os.path.join(self.profile, "round-trip.docx")
        url = uno.systemPathToFileUrl(path)
        cursor = self.doc.Text.createTextCursor()
        self.doc.Text.insertString(cursor, "DOCX title", False)
        self.doc.Text.insertControlCharacter(cursor, uno.getConstantByName("com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK"), False)
        self.doc.Text.insertString(cursor, "DOCX body.", False)

        filter_name = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        filter_name.Name = "FilterName"
        filter_name.Value = "Office Open XML Text"
        self.doc.storeAsURL(url, (filter_name,))
        self.doc.close(True)
        self.doc = self.desktop.loadComponentFromURL(url, "_blank", 0, ())
        self.formatter = main.Format(self.context, self.doc)
        self.formatter.format_document_structure({"infer": True, "title": {"bold": True}})
        paragraphs = list(self.formatter._paragraphs())
        self.assertEqual(paragraphs[0].String, "DOCX title")
        self.assertEqual(paragraphs[0].ParaStyleName, "Title")

    def test_multiple_table_selection(self):
        first = self.doc.createInstance("com.sun.star.text.TextTable")
        first.initialize(1, 1)
        self.doc.Text.insertTextContent(self.doc.Text.createTextCursor(), first, False)
        cursor = self.doc.Text.createTextCursor()
        cursor.gotoEnd(False)
        second = self.doc.createInstance("com.sun.star.text.TextTable")
        second.initialize(1, 1)
        self.doc.Text.insertTextContent(cursor, second, False)

        self.assertEqual(len(self.formatter.get_tables("table_1")), 1)
        self.assertEqual(len(self.formatter.get_tables("table_2")), 1)
        self.formatter.format_table(self.formatter.get_tables("table_2")[0], {"table_background": "FF0000"})
        self.assertNotEqual(first.getCellByName("A1").BackColor, 0xFF0000)
        self.assertEqual(second.getCellByName("A1").BackColor, 0xFF0000)


if __name__ == "__main__":
    unittest.main()

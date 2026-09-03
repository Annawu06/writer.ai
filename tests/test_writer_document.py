import os
import shutil
import subprocess
import tempfile
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

    def test_table_formatting(self):
        table = self.doc.createInstance("com.sun.star.text.TextTable")
        table.initialize(2, 2)
        self.doc.Text.insertTextContent(self.doc.Text.createTextCursor(), table, False)
        table.getCellByName("A1").String = "Header"
        table.getCellByName("B1").String = "Value"
        table.getCellByName("A2").String = "Row"

        self.formatter.format_table(table, {
            "header": True,
            "header_background": "0ABAB5",
            "first_column_bold": True,
            "zebra": True,
            "zebra_background": "F2F2F2",
            "align": "center",
        })
        self.assertEqual(table.getCellByName("A1").BackColor, 0x0ABAB5)
        self.assertEqual(table.getCellByName("A2").BackColor, 0xF2F2F2)
        self.assertEqual(int(table.getCellByName("A1").createTextCursor().ParaAdjust), 3)

    def test_empty_document(self):
        paragraphs = list(self.formatter._paragraphs())
        self.assertTrue(all(not paragraph.String for paragraph in paragraphs))


if __name__ == "__main__":
    unittest.main()

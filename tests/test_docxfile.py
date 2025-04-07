import unittest

from src.docx2md.docxfile import DocxFileError, DocxFile

class TestDocxFile(unittest.TestCase):

    def test_file_not_found(self):
        with self.assertRaises(FileNotFoundError):
            DocxFile("notexist.txt")

    def test_not_zip_file(self):
        with self.assertRaises(DocxFileError):
            DocxFile("tests/data/dummy.txt")

    def test_not_dodx_file(self):
        with self.assertRaises(DocxFileError):
            DocxFile("tests/data/dummy.zip")

    def test_valid_docx_file(self):
        docx = DocxFile("tests/data/paragraph.docx")
        xml = docx.document()
        self.assertTrue(b"<w:document" in xml)

        docx.close()
        with self.assertRaises(FileNotFoundError):
            xml = docx.document()


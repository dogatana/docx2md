import unittest
import os.path

from src.docx2md.converter import Converter
from src.docx2md.docxfile import DocxFile
from src.docx2md.docxmedia import DocxMedia

class TestDocxFile(unittest.TestCase):

    def test_convert_paragrah(self):
        self.diff("paragraph")

    def test_convert_list(self):
        self.diff("list")

    def test_convert_heading(self):
        self.diff("heading")
    
    def test_convert_table(self):
        self.diff("table")

    def test_convert_image(self):
        self.diff("image")

    def diff(self, name):
        file = os.path.join("tests/data", name + ".docx")

        target_dir = os.path.splitext(file)[0]

        docx = DocxFile(file)
        media = DocxMedia(docx)
        md_result = self.convert(docx, target_dir, media, False)
        md_text = self.read_md(target_dir)

        self.assertEqual(md_result, md_text)
    
    def convert(self, docx, target_dir, media, use_md_table):
        xml_text = docx.document()
        converter = Converter(xml_text, media, use_md_table)
        return converter.convert()

    def read_md(self, target_dir):
        file = os.path.join(target_dir, "README.md")
        with open(file, encoding="utf-8") as f:
            return f.read()
import unittest
import os.path

import docx2md

class TestDocxFile(unittest.TestCase):

    def test_convert_paragrah(self):
        self.diff("paragraph")

    def test_convert_heading(self):
        self.diff("heading")
    
    def test_convert_table(self):
        self.diff("table")
    
    def diff(self, name):
        file = os.path.join("tests/data", name + ".docx")

        target_dir = os.path.splitext(file)[0]

        docx = docx2md.DocxFile(file)
        media = docx2md.DocxMedia(docx)
        md_result = docx2md.convert(docx, target_dir, media, False)
        md_text = self.read_md(target_dir)

        self.assertEqual(md_result, md_text)
    
    def read_md(self, target_dir):
        file = os.path.join(target_dir, "README.md")
        with open(file, encoding="utf-8") as f:
            return f.read()
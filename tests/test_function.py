import unittest
import glob
import os.path

import docx2md

class TestDocxFile(unittest.TestCase):

    def test_convert(self):
        for file in glob.glob("tests/data/*.docx"):
            print(f"# {file}")
            target_dir = os.path.splitext(file)[0]
    
            md_result = docx2md.convert(file, target_dir, save_images=False, debug=False)
            md_text = self.read_md(target_dir)

            self.assertEqual(md_result, md_text)

    
    def read_md(self, target_dir):
        file = os.path.join(target_dir, "README.md")
        with open(file, encoding="utf-8") as f:
            return f.read()
import unittest

from docx2md.docxmedia import Media


class TestMedia(unittest.TestCase):
    def test_png_media(self):
        m = Media("id", "test.png")
        self.assertFalse(m.use_alt)

    def test_xxx_media(self):
        m = Media("id", "test.xxx")
        self.assertTrue(m.use_alt)
        self.assertEqual(m.alt_path, "test.png")

import unittest

from src.docx2md.docxmedia import DocxMedia


class FakeDocxFile:
    def __init__(self, text):
        self.text = text

    def rels(self):
        return (
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships>'
            + self.text
            + b"</Relationships>"
        )


class TestDocxMedia(unittest.TestCase):
    def test_length_is_zoro(self):
        """no definition"""
        docx = FakeDocxFile(b"")
        m = DocxMedia(docx)
        self.assertEqual(len(m), 0)
        self.assertFalse("id" in m)
        self.assertIsNone(m["id"])

    def test_no_media(self):
        """no media"""
        docx = FakeDocxFile(b'<Relationship Id="id" Target="test.png"/>')
        m = DocxMedia(docx)
        self.assertEqual(len(m), 0)
        self.assertFalse("id" in m)
        self.assertIsNone(m["id"])

    def test_one_png_media(self):
        """media/test.png"""
        docx = FakeDocxFile(b'<Relationship Id="id" Target="media/test.png"/>')
        m = DocxMedia(docx)
        self.assertEqual(len(m), 1)
        self.assertTrue("id" in m)

        self.assertEqual(m["id"].path, "media/test.png")
        self.assertFalse(m["id"].use_alt)

    def test_one_xxx_media(self):
        """media/test.xxx"""
        docx = FakeDocxFile(b'<Relationship Id="id" Target="media/test.xxx"/>')
        m = DocxMedia(docx)
        self.assertEqual(len(m), 1)
        self.assertTrue("id" in m)

        self.assertEqual(m["id"].path, "media/test.xxx")
        self.assertEqual(m["id"].alt_path, "media/test.png")
        self.assertTrue(m["id"].use_alt)

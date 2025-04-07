import unittest

from src.docx2md.converter import Converter
from src.docx2md.docxmedia import Media


def build_xml(text):
    head = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><document><body>'
    tail = b"</body></document>"
    return head + text + tail


class TestConvert(unittest.TestCase):
    def test_no_blip(self):
        xml = build_xml(
            b"""
        <p>
          <r><t>image</t></r>
        </p>"""
        )
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "image")

    def test_invalid_id(self):
        xml = build_xml(
            b"""
        <p>
          <r><t>image</t></r>
          <r><drawing><blip embed="id"/></drawing></r>
        </p>"""
        )
        converter = Converter(xml, {}, False)
        with self.assertRaises(KeyError):
            _ = converter.convert()

    def test_png(self):
        media = dict(id=Media("id", "media/image.png"))
        xml = build_xml(
            b"""
        <p>
          <r><t>image</t></r>
          <r><drawing><blip embed="id"/></drawing></r>
        </p>"""
        )
        converter = Converter(xml, media, False)
        result = converter.convert()
        self.assertEqual(result, 'image<img src="media/image.png" id="image1">')

    def test_png_and_enf(self):
        media = dict(
            id1=Media("id1", "media/image1.png"),
            id2=Media("id2", "media/image2.enf"),
        )
        xml = build_xml(
            b"""
        <p>
          <r><t>image</t></r>
          <r><drawing><blip embed="id1"/></drawing></r>
          <r><drawing><blip embed="id2"/></drawing></r>
        </p>"""
        )
        converter = Converter(xml, media, False)
        result = converter.convert()
        self.assertEqual(
            result,
            'image<img src="media/image1.png" id="image1">'
            + '<img src="media/image2.png" id="image2">',
        )

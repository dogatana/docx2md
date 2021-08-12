import unittest

from docx2md import Converter


def build_xml(text):
    head = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><document><body>'
    tail = b"</body></document>"
    return head + text + tail

class TestConvert(unittest.TestCase):
    def test_ignore_toc(self):
        xml = build_xml(b"""
        <sdt>
          <p><t>table of contents</t></p>
        </sdt>
        <p><t>text2</t></p>""")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "text2")

    def test_heading_123(self):
        xml = build_xml(b"""
        <p>
          <pPr><pStyle val="1"/></pPr>
          <t>h1</t>
        </p>
        <p><t>text1</t></p>
        <p>
          <pPr><pStyle val="2"/></pPr>
          <t>h2</t>
        </p>
        <p><t>text2</t></p>
        <p>
          <pPr><pStyle val="3"/></pPr>
          <t>h3</t>
        </p>
        <p><t>text3</t></p>""")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "\n\n".join(["# h1", "text1", "## h2", "text2", "### h3", "text3"]))
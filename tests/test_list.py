import unittest

from docx2md import Converter


def build_xml(text):
    head = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><document><body>'
    tail = b"</body></document>"
    return head + text + tail

class TestConvert(unittest.TestCase):
    def test_ul(self):
        xml = build_xml(b"""
        <p>
          <pPr>
            <pStyle val="a"/>
            <numPr>
              <ilvl val="0"/>
            </numPr>
          </pPr>
          <t>text1</t>
        </p>
        <p>
          <pPr>
            <pStyle val="a"/>
            <numPr>
              <ilvl val="0"/>
            </numPr>
          </pPr>
          <t>text2</t>
        </p>""")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "* text1\n* text2")

    def test_nested_ul(self):
        xml = build_xml(b"""
        <p>
          <pPr>
            <pStyle val="a"/>
            <numPr>
              <ilvl val="0"/>
            </numPr>
          </pPr>
          <t>text1</t>
        </p>
        <p>
          <pPr>
            <pStyle val="a"/>
            <numPr>
              <ilvl val="1"/>
            </numPr>
          </pPr>
          <t>text2</t>
        </p>""")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "* text1\n    * text2")

    def test_ul_and_paragraphs(self):
        xml = build_xml(b"""
        <p><t>before</t></p>
        <p>
          <pPr>
            <pStyle val="a"/>
            <numPr>
              <ilvl val="0"/>
            </numPr>
          </pPr>
          <t>text1</t>
        </p>
        <p><t>after</t></p>""")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "\n\n".join(["before", "* text1", "after"]))
import unittest

from docx2md import Converter


def build_xml(text):
    head = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><document>'
    tail = b"</document>"
    return head + text + tail

class TestConvert(unittest.TestCase):

    def test_no_body(self):
        """ no body tag """
        xml = build_xml(b"")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "")

    def test_blank_body(self):
        """ blank body """
        xml = build_xml(b"<body></body>")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "")

    def test_one_paragraph(self):
        xml = build_xml(b"<body><p> <t> text </t> </p> </body>")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "text")

    def test_one_paragraph_with_break(self):
        xml = build_xml(b"<body><p> <t> text1 </t> <br/> <t> text2 </t> </p> </body>")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "text1 <br> text2")

    def test_two_paragraphs(self):
        xml = build_xml(b"""<body>
        <p><t>text1</t></p>
        <p><t>text2</t></p>
        </body>""")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "text1\n\ntext2")

    def test_two_paragraphs_with_blank_lines(self):
        xml = build_xml(b"""<body>
        <p><t>text1</t></p>
        <p></p>
        <p></p>
        <p></p>
        <p></p>
        <p><t>text2</t></p>
        </body>""")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, "text1\n\ntext2")

    def test_two_paragraphs_with_page_break(self):
        xml = build_xml(b"""<body>
        <p><t>text1</t></p>
        <br type="page"/>
        <p><t>text2</t></p>
        </body>""")
        converter = Converter(xml, {}, False)
        result = converter.convert()
        self.assertEqual(result, 'text1\n\n<div class="break"></div>\n\ntext2')
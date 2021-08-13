import unittest

from docx2md import Converter


def build_xml(text):
    head = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><document><body>'
    tail = b"</body></document>"
    return head + text + tail


class TestConvert(unittest.TestCase):
    def test_3by3_table(self):
        xml = build_xml(b"""
          <tbl>
            <tr><tc><p><t>a</t></p></tc><tc><p><t>b</t></p></tc><tc><p><t>c</t></p></tc></tr>
            <tr><tc><p><t>d</t></p></tc><tc><p><t>e</t></p></tc><tc><p><t>f</t></p></tc></tr>
            <tr><tc><p><t>g</t></p></tc><tc><p><t>h</t></p></tc><tc><p><t>i</t></p></tc></tr>
          </tbl>
        """)

        converter = Converter(xml, {}, False)
        result = converter.convert()
        expected = "\n".join(
            ['<table id="table1">'] +
            ["<tr>"] + ([f"<td>{x}</td>" for x in "abc"]) +  ["</tr>"] +
            ["<tr>"] + ([f"<td>{x}</td>" for x in "def"]) +  ["</tr>"] +
            ["<tr>"] + ([f"<td>{x}</td>" for x in "ghi"]) +  ["</tr>"] +
            ["</table>"])
        self.assertEqual(result, expected)

    def test_3by3_md_table(self):
        xml = build_xml(b"""
          <tbl>
            <tr><tc><p><t>a</t></p></tc><tc><p><t>b</t></p></tc><tc><p><t>c</t></p></tc></tr>
            <tr><tc><p><t>d</t></p></tc><tc><p><t>e</t></p></tc><tc><p><t>f</t></p></tc></tr>
            <tr><tc><p><t>g</t></p></tc><tc><p><t>h</t></p></tc><tc><p><t>i</t></p></tc></tr>
          </tbl>
        """)

        converter = Converter(xml, {}, True)
        result = converter.convert()
        expected = "\n".join([
            "| # | # | # |",
            "|---|---|---|",
            "|a|b|c|",
            "|d|e|f|",
            "|g|h|i|",])
        self.assertEqual(result, expected)

    def test_merged_table(self):
        xml = build_xml(b"""
          <tbl>
            <tr>
               <tc>
                 <gridSpan val="2"/>
                 <vMerge val="restart"/>
                 <p><t>a</t></p>
                 <p><t>b</t></p>
                 <p><t>d</t></p>
                 <p><t>e</t></p>
               </tc>
               <tc>
                 <p><t>c</t></p>
               </tc>
            </tr>
            <tr>
               <tc>
                 <gridSpan val="2"/>
                 <vMerge/>
               </tc>
               <tc>
                 <p><t>f</t></p>
               </tc>
            </tr>
            <tr>
                <tc><p><t>g</t></p></tc>
                <tc><p><t>h</t></p></tc>
                <tc><p><t>i</t></p></tc>
            </tr>
          </tbl>
        """)

        converter = Converter(xml, {}, False)
        result = converter.convert()
        expected = """<table id="table1">
<tr>
<td colspan="2" rowspan="2">a<br>b<br>d<br>e</td>
<td>c</td>
</tr>
<tr>
<td>f</td>
</tr>
<tr>
<td>g</td>
<td>h</td>
<td>i</td>
</tr>
</table>"""
        self.assertEqual(result, expected)

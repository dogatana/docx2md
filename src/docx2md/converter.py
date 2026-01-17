import collections
import io
import re

from lxml import etree

from . import utils


class Converter:
    def __init__(self, xml_text, media, use_md_table):
        self.tree = etree.fromstring(xml_text)
        utils.strip_ns_prefix(self.tree)
        self.media = media
        self.image_counter = self.counter()
        self.table_counter = self.counter()
        self.use_md_table = use_md_table

    def counter(self, start=1):
        count = start - 1

        def inc():
            nonlocal count
            count += 1
            return count

        return inc

    def convert(self):
        self.in_list = False

        of = io.StringIO()
        body = self.get_first_element(self.tree, "//body")
        self.parse_node(of, body)

        return re.sub(r"\n{2,}", "\n\n", of.getvalue()).strip()

    def get_first_element(self, node, xpath):
        tags = node.xpath(xpath)
        return tags[0] if len(tags) > 0 else None

    def get_sub_text(self, node):
        of = io.StringIO()
        self.parse_node(of, node)
        return of.getvalue().strip()

    # {'sectPr', 'tbl', 'bookmarkEnd', 'moveToRangeEnd', 'bookmarkStart', 'commentRangeStart', 'moveFromRangeEnd', 'sdt', 'p'}
    BODY_IGNORE = [
        "sectPr",
        # "tbl",
        # "p"}
        "bookmarkStart",
        "bookmarkEnd",
        "moveToRangeEnd",
        "commentRangeStart",
        "moveFromRangeEnd",
        "sdt",
    ]

    def parse_node(self, of, node):
        for child in node.getchildren():
            tag_name = child.tag
            if tag_name in self.BODY_IGNORE:
                continue
            elif tag_name == "p":
                self.parse_p(of, child)
            elif tag_name == "tbl":
                self.parse_tbl(of, child)
            else:
                print("# ** skip", tag_name)

    P_IGNORE = [
        # "ins", "r",
        "pPr",
        "sdt",
        "moveFrom",
        "moveTo",
        "hyperlink",
        "del",
        "proofErr",
        "moveToRangeStart",
        "commentRangeStart",
        "commentRangeEnd",
        "bookmarkStart",
        "bookmarkEnd",
        "moveFromRangeStart",
        "moveFromRangeEnd",
    ]

    def parse_p(self, of, node):
        def out_p(text):
            print("", file=of)
            print(text, file=of)
            print("", file=of)

        sub_text = self.parse_p_text(node)

        # fix issue #6
        if node.getparent().tag != "body":
            out_p(sub_text)
            return

        pStyle = self.get_first_element(node, "./pPr/pStyle")
        if pStyle is None:
            if self.in_list:
                self.in_list = False
            out_p(sub_text)
            return

        style = pStyle.attrib["val"]
        # im having issues with this as style comes as "Heading<digit>"
        # maybe not pure MS-docx but google-worksuite generated docx?
        if "Heading" in style:
            style = style.replace("Heading", "")
        if style.isdigit():
            print("#" * (int(style)), sub_text, file=of)
        elif style[0] == "a":
            ilvl = self.get_first_element(node, ".//ilvl")
            if ilvl is None:
                return
            level = int(ilvl.attrib["val"])
            print("    " * level + "*", sub_text, file=of)
        else:
            out_p(sub_text)

    def parse_p_text(self, node):
        of = io.StringIO()
        for r in node.xpath("./r|./ins/r"):
            self.parse_r(of, r)
        return of.getvalue()

    R_IGNORE = [
        # "pict", "t", "br", "drawing",
        "tab",
        "lastRenderedPageBreak",
        "rPr",
        "instrText",
        "delText",
        "fldChar",
    ]

    def parse_r(self, of, node):
        for child in node.getchildren():
            tag_name = child.tag
            if tag_name == "t":
                text = child.text or " "
                text = text.replace("\u00a0", "&nbsp;")
                print(text, end="", file=of)
            elif tag_name == "br":
                if child.attrib.get("type") == "page":
                    print('<div class="page"></div>', file=of)
                else:
                    print("<br>", end="", file=of)
            elif tag_name == "drawing":
                blip = self.get_first_element(child, ".//blip")
                if blip is None:
                    print("[parse_r]", tag_name)
                    continue
                id = blip.attrib.get("embed")
                if id is None:
                    print("[parse_r]", tag_name)
                    continue
                self.emit_image(of, id)
            elif tag_name == "object":
                imagedata = self.get_first_element(child, ".//imagedata")
                if imagedata is None:
                    print("[parse_r]", tag_name)
                    continue
                id = imagedata.attrib.get("id")
                if id is None:
                    print("[parse_r]", tag_name)
                    continue
                self.emit_image(of, id)
            elif tag_name == "pict":
                print("<{tag_name}>", end="", file=of)
            elif tag_name in self.R_IGNORE:
                continue
            else:
                print("[parse_r]", tag_name)

    def parse_tbl(self, of, node):
        properties = self.get_table_properties(node)
        if self.use_md_table:
            self.emit_md_table(of, node, len(properties[0]))
        else:
            self.emit_html_table(of, node, properties)

    def emit_md_table(self, of, node, col_size):
        print("", file=of)
        print("| # " * (col_size) + "|", file=of)
        print("|---" * col_size + "|", file=of)
        for tag_tr in node.xpath(".//tr"):
            print("|", end="", file=of)
            for tag_tc in tag_tr.xpath(".//tc"):
                span = 1
                gridSpan = self.get_first_element(tag_tc, ".//gridSpan")
                if gridSpan is not None:
                    span = int(gridSpan.attrib["val"])
                sub_text = self.get_sub_text(tag_tc)
                text = re.sub(r"\n+", "<br>", sub_text)
                print(text, end="", file=of)
                print("|" * span, end="", file=of)
            gridAfter = self.get_first_element(tag_tr, ".//gridAfter")
            if gridAfter is not None:
                val = int(gridAfter.attrib["val"])
                print("|" * val, end="", file=of)
            print("", file=of)
        print("", file=of)

    def emit_html_table(self, of, node, properties):
        id = f"table{self.table_counter()}"
        print(f'\n<table id="{id}">', file=of)
        for y, tr in enumerate(node.xpath(".//tr")):
            print("<tr>", file=of)
            x = 0
            for tc in tr.xpath(".//tc"):
                prop = properties[y][x]
                colspan = prop.span
                attr = "" if colspan <= 1 else f' colspan="{colspan}"'
                rowspan = prop.merge_count
                attr += "" if rowspan == 0 else f' rowspan="{rowspan}"'

                sub_text = self.get_sub_text(tc)
                text = re.sub(r"\n+", "<br>", sub_text)
                if not prop.merged or prop.merge_count != 0:
                    print(f"<td{attr}>{text}</td>", file=of)
                x += colspan
            gridAfter = self.get_first_element(tr, ".//gridAfter")
            if gridAfter is not None:
                val = int(gridAfter.attrib["val"])
                for _ in range(val):
                    print("<td></td>", file=of)
            print("</tr>", file=of)
        print("</table>", file=of)

    def get_table_properties(self, node):
        CellProperty = collections.namedtuple(
            "CellProperty", ["span", "merged", "merge_count"]
        )
        properties = []
        for tr in node.xpath(".//tr"):
            row_properties = []
            for tc in tr.xpath(".//tc"):
                span = 1
                gridSpan = self.get_first_element(tc, ".//gridSpan")
                if gridSpan is not None:
                    span = int(gridSpan.attrib["val"])
                merged = False
                merge_count = 0
                vMerge = self.get_first_element(tc, ".//vMerge")
                if vMerge is not None:
                    merged = True
                    val = vMerge.attrib.get("val")
                    merge_count = 1 if val == "restart" else 0
                prop = CellProperty(span, merged, merge_count)
                row_properties.append(prop)
                for _ in range(span - 1):
                    row_properties.append(
                        CellProperty(0, prop.merged, prop.merge_count)
                    )
            gridAfter = self.get_first_element(tr, ".//gridAfter")
            if gridAfter is not None:
                val = int(gridAfter.attrib["val"])
                for _ in range(val):
                    row_properties.append(CellProperty(1, False, 0))
            properties.append(row_properties)

        for y in range(len(properties) - 1):
            for x in range(len(properties[0])):
                prop = properties[y][x]
                if prop.merge_count > 0:
                    count = 0
                    for ynext in range(y + 1, len(properties)):
                        cell = properties[ynext][x]
                        if cell.merged and cell.merge_count == 0:
                            count += 1
                        elif not cell.merged or cell.merge_count > 0:
                            break
                    properties[y][x] = CellProperty(
                        prop.span, prop.merged, prop.merge_count + count
                    )
        return properties

    def parse_drawing(self, of, node):
        """pictures"""
        blip = self.get_first_element(node, ".//blip")
        if blip is None:
            return

        embed_id = blip.attrib.get("embed")
        if embed_id is None or embed_id not in self.media:
            return

        tag_id = f"image{self.image_counter()}"
        print(
            f'<img src="{self.media[embed_id].alt_path}" id="{tag_id}">',
            end="",
            file=of,
        )

    def emit_image(self, of, id):
        tag_id = f"image{self.image_counter()}"
        print(f'<img src="{self.media[id].alt_path}" id="{tag_id}">', end="", file=of)


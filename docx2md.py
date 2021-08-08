import sys
import os.path
import re
import base64
import io
import re

import zipfile

from lxml import etree


class DocxFileError(RuntimeError):
    pass

class DocxFile:
    def __init__(self, filename):
        if not os.path.isfile(filename):
            raise FileNotFoundError(f"{filename} not found")
        if not zipfile.is_zipfile(filename):
            raise DocxFileError(f"{filename} is not a zip file")

        self.docx = zipfile.ZipFile(filename)
        if "word/document.xml" not in self.namelist():
            raise DocxFileError(f"{filename} does not include word/document.xml")

    def document(self):
        if self.docx is None:
            raise FileNotFoundError
        return self.docx.read("word/document.xml")

    def extract_images(self, target_dir):
        images = self.__images()
        if not images:
            return

        if not os.path.exists(target_dir):
            os.makedirs(target_dir)

        for image in images:
            file = os.path.join(target_dir, image[len("word/media/") :])
            with open(file, "wb") as f:
                f.write(self.read(image))

    def __images(self):
        return [f for f in self.namelist() if f.startswith("word/media/")]

    def namelist(self):
        if self.docx is None:
            raise FileNotFoundError
        return self.docx.namelist()

    def read(self, filename):
        if self.docx is None:
            raise FileNotFoundError
        if filename not in self.namelist():
            return None
        return self.docx.read(filename)
    
    def close(self):
        self.docx.close()
        self.docx = None

def parse_docx(file):
    docx = DocxFile(file)
    save_xml(file, docx.document())

    tree = etree.fromstring(docx.document())
    strip_ns_prefix(tree)
    # resolver = NamespaceResolver(doc.nsmap)

    body = get_first_element(tree, "//body")

    of = io.StringIO()
    parse_node(of, body)
    text = re.sub(r"\n{2,}", "\n\n", of.getvalue()).strip()
    print("-" * 10, text, "-" * 10, sep="\n")
    open("out.md", "w", encoding="utf-8").write(text)


def save_xml(file, text):
    # save for debug
    xml_file = os.path.splitext(file)[0] + ".xml"
    if os.path.exists(xml_file):
        return
    open(xml_file, "wb").write(text)
    print("save to", xml_file)


def strip_ns_prefix(tree):
    # xpath query for selecting all element nodes in namespace
    query = "descendant-or-self::*[namespace-uri()!='']"
    # for each element returned by the above xpath query...
    for element in tree.xpath(query):
        # replace element name with its local name
        element.tag = etree.QName(element).localname
    return tree


STOP_PARSING = [
    "pPr",
    "rPr",
    "sectPr",
    "pStyle",
    "wPr",
    "numPr",
    "ind",
    "shapetype",
    "tab",
    "sdt"
]
NOT_PROCESS = ["shape", "txbxContent"]


def parse_node(of, node):
    if node.tag in STOP_PARSING:
        return
    for child in node.getchildren():
        if child is None:
            print("** none")
            continue
        tag = child.tag
        if tag in STOP_PARSING:
            continue
        if tag == "p":
            parse_p(of, child)
        elif tag == "r":
            parse_node(of, child)
        elif tag == "br":
            attr = ' class="page"'
            if get_attr(child, "type") == "page":
                print('\n<div class="break"></div>\n', file=of)
            else:
                print("<br>", end="", file=of)
        elif tag == "t":
            print(child.text or " ", end="", file=of)
        elif tag == "drawing":
            parse_drawing(of, child)
        elif tag == "textbox":
            # print("\n\n--\n")
            parse_node(of, child)
            # print("\n\n--\n")
        elif tag == "tbl":
            parse_tbl(of, child)
            # print("\n<table>", file=of)
            # parse_node(of, child)
            # print("</table>", file=of)
        elif tag == "tr":
            print("<tr>", file=of)
            parse_node(of, child)
            print("</tr>", file=of)
        elif tag == "tc":
            parse_tc(of, child)
        else:
            parse_node(of, child)


def parse_tbl(of, node):
    properties = get_table_properties(node)
    print("\n<table>", file=of)
    for y, tag_tr in enumerate(node.xpath(".//tr")):
        print("<tr>", file=of)
        x = 0
        for tag_tc in tag_tr.xpath(".//tc"):
            prop = properties[y][x]
            colspan = prop["span"]
            attr = "" if colspan <= 1 else f' colspan="{colspan}"'
            rowspan = prop["merge_count"]
            attr += "" if rowspan == 0 else f' rowspan="{rowspan}"'

            sub_text = get_sub_text(tag_tc)
            text = re.sub(r"\n+", "<br>", sub_text)
            if not prop["merged"] or prop["merge_count"] != 0:
                print(f"<td{attr}>{text}</td>", file=of)
            x += colspan
        gridAfter =get_first_element(tag_tr, ".//gridAfter")
        if gridAfter is not None:
            val = len(get_attr(gridAfter, "val"))
            for _ in range(val):
                print("<td></td>", file=of)
        print("</tr>", file=of)
    print("</table>", file=of)

def get_table_properties(node):
    properties = []
    for tag_tr in node.xpath(".//tr"):
        row_property = []
        for tag_tc in tag_tr.xpath(".//tc"):
            span = 1
            gridSpan = get_first_element(tag_tc, ".//gridSpan")
            if gridSpan is not None:
                span = int(get_attr(gridSpan, "val"))
            merged = False
            merge_count = 0
            vMerge = get_first_element(tag_tc, ".//vMerge")
            if vMerge is not None:
                merged = True
                val = get_attr(vMerge, "val")
                merge_count = 1 if val == "restart" else 0
            prop = { 
                "span": span, 
                "merged": merged, 
                "merge_count": merge_count
            }
            row_property.append(prop)
            copied_prop = prop.copy()
            copied_prop["span"] = 0
            for _ in range(span - 1):
                row_property.append(copied_prop)
        gridAfter =get_first_element(tag_tr, ".//gridAfter")
        if gridAfter is not None:
            val = len(get_attr(gridAfter, "val"))
            for _ in range(val):
                row_property.append({"span": 1, "merged": False, "merge_count": 0})
        properties.append(row_property)
    print("y", len(properties))
    print("x", len(properties[0]))
    print(properties)
    print(get_sub_text(node))
    for y in range(len(properties) - 1):
        for x in range(len(properties[0])):
            print(y, x)
            if properties[y][x]["merge_count"] > 0:
                count = 0
                for ynext in range(y + 1, len(properties)):
                    cell = properties[ynext][x]
                    if cell["merged"] and cell["merge_count"] == 0:
                        count += 1
                    elif not cell["merged"] or cell["merge_count"] > 0:
                        break
                properties[y][x]["merge_count"] += count
    return properties
    
    for y in range(len(properties)):
        for x in range(len(properties[0])):
            print(y, x, properties[y][x])
    


            




def parse_tc(of, node):
    sub_text = get_sub_text(node)
    text = re.sub(r"\n+", "<br>", sub_text)

    attr = ""
    span = get_first_element(node, ".//gridSpan")
    if span is not None:
        val = get_attr(span, "val")
        attr = f' colspan="{val}"'
    print(f"<td{attr}>", end="", file=of)
    print(text, end="", file=of)
    print("</td>", file=of)


def get_sub_text(node):
    of = io.StringIO();
    parse_node(of, node)
    return of.getvalue().strip()

in_list = False


def parse_p(of, node):
    global in_list
    """ parse paragraph """
    pStyle = get_first_element(node, ".//pStyle")
    if pStyle is None:
        if in_list:
            in_list = False
        print("", file=of)
        parse_node(of, node)
        print("", file=of)
        return

    sub_of = io.StringIO()
    parse_node(sub_of, node)
    sub_text = sub_of.getvalue().strip()
    if not sub_text:
        return

    if not in_list:
        print("", file=of)
        in_list = True

    style = get_attr(pStyle, "val")
    if style.isdigit():
        print("#" * (int(style)), sub_text, file=of)
    elif style[0] == "a":
        ilvl = get_first_element(node, ".//ilvl")
        if ilvl is None:
            return
        level = int(get_attr(ilvl, "val"))
        print("    " * level + "*", sub_text, file=of)
    else:
        raise RuntimeError("pStyle: " + style)


def get_first_element(tree, xpath):
    tags = tree.xpath(xpath)
    return tags[0] if len(tags) > 0 else None


def get_attr(tag, name):
    for key, value in tag.attrib.items():
        if key.endswith("}" + name):
            return value
    return None


def parse_drawing(of, node):
    tag = get_first_element(node, ".//cNvPr")
    if tag is not None and "id" in tag.attrib:
        id = tag.attrib["id"]
        print(f'<img src="image{id}.png">', file=of)
    else:
        print(f"*** no pictures")


def decode_shape(gfxdata, file):
    decoded = base64.b64decode(gfxdata.replace("&#xA;", "\n"))
    open(file, "wb").write(decoded)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage: python docx2md <word.docx>")
        exit()
    file = sys.argv[1]
    if not os.path.exists(file):
        print(file, "does not exit.")
        exit()
    parse_docx(file)

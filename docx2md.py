import sys
import os.path
import re
import base64
import io
import re

import zipfile

from lxml import etree


class InvalidDocxFileError(RuntimeError):
    pass

class DocxFile:
    def __init__(self, filename):
        if not os.path.isfile(filename):
            raise FileNotFoundError()
        if not zipfile.is_zipfile(filename):
            raise InvalidDocxFileError()

        self.docx = zipfile.ZipFile(filename)
        if "word/document.xml" not in self.namelist():
            raise InvalidDocxFileError()

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


class NamespaceResolver:
    def __init__(self, nsmap):
        self.map = {}
        for k, v in nsmap.items():
            self.map[v] = k

    def tag(self, name):
        m = re.search(r"\{(.*?)\}(.*)", name)
        return f"{self.map[m.group(1)]}:{m.group(2)}"


def parse_docx(file):
    docx = DocxFile(file)
    save_xml(file, docx.document())

    tree = etree.fromstring(docx.document())
    strip_ns_prefix(tree)
    # resolver = NamespaceResolver(doc.nsmap)

    body = get_first_element(tree, "//body")

    of = io.StringIO()
    parse_node(of, body, 0)
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


def parse_node(of, node, depth):
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
            parse_p(of, child, depth + 1)
        elif tag == "r":
            parse_node(of, child, depth + 1)
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
            parse_node(of, child, depth + 1)
            # print("\n\n--\n")
        elif tag == "tbl":
            print("\n<table>", file=of)
            parse_node(of, child, depth + 1)
            print("</table>", file=of)
        elif tag == "tr":
            print("<tr>", file=of)
            parse_node(of, child, depth + 1)
            print("</tr>", file=of)
        elif tag == "tc":
            print("<td>", end="", file=of)
            parse_tc(of, child, depth + 1)
            print("</td>", file=of)
        else:
            if tag not in NOT_PROCESS:
                print("#", tag)
            parse_node(of, child, depth + 1)


def parse_tc(of, node, depth):
    sub_of = io.StringIO()
    parse_node(sub_of, node, depth + 1)
    sub_text = sub_of.getvalue().strip()
    text = re.sub(r"\n+", "<br>", sub_text)
    print(text, end="", file=of)


in_list = False


def parse_p(of, node, depth):
    global in_list
    """ parse paragraph """
    pStyle = get_first_element(node, ".//pStyle")
    if pStyle is None:
        if in_list:
            in_list = False
        print("", file=of)
        parse_node(of, node, depth + 1)
        print("", file=of)
        return

    sub_of = io.StringIO()
    parse_node(sub_of, node, depth + 1)
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

import sys
import os.path
import re
import base64
import io
import zipfile

from lxml import etree

class InvalidDocxfileError(RuntimeError):
    pass

class DocxFile:
    def __init__(self, filename):
        if not os.path.isfile(filename):
            raise FileNotFoundError()
        if not zipfile.is_zipfile(filename):
            raise InvalidDocxfileError()

        self.docx = zipfile.ZipFile(filename)
        if "word/document.xml" not in self.namelist():
            raise InvalidDocxfileError()
    
    def document(self):
        return self.docx.read("word/document.xml")

    def extract_images(self, target_dir):
        images = self.__images()
        if not images:
            return

        if not os.path.exists(target_dir):
            os.makedirs(target_dir)

        for image in images:
            file = os.path.join(target_dir, image[len("word/media/"):])
            with open(file, "wb") as f:
                f.write(self.read(image))

    def __images(self):
        return [f for f in self.namelist() if f.startswith("word/media/")]

    def namelist(self):
        return self.docx.namelist()
    
    def read(self, filename):
        if filename not in self.namelist():
            return None
        return self.docx.read(filename)
    
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
    doc = etree.fromstring(docx.document())
    strip_ns_prefix(doc)
    # resolver = NamespaceResolver(doc.nsmap)

    body = doc.xpath("//body")[0]
    parse_tag(body, 0)

def strip_ns_prefix(tree):
    #xpath query for selecting all element nodes in namespace
    query = "descendant-or-self::*[namespace-uri()!='']"
    #for each element returned by the above xpath query...
    for element in tree.xpath(query):
        #replace element name with its local name
        element.tag = etree.QName(element).localname
    return tree

IGNORE = [
    "pPr", "rPr", "rFonts", "sectPr", "pgSz", "pgMar", "cols", "docGrid", 
    "shapetype", "stroke", "path",
    "pict", "textbox", "txbxContent"
    "spacing", "ind", "b"
]
SHOW = [
    "shape"
]
def parse_tag(root, depth):
    for child in root.getchildren():
        if  child is None:
            print("** none")
            continue
        tag = child.tag
        if tag == "p":
            parse_tag(child, depth + 1)
            print("\n")
        elif tag == "r":
            parse_tag(child, depth + 1)
        elif tag == "br":
            print("<br>")
        elif tag == "t":
            print(child.text or " ", end="")
        elif tag == "drawing":
            parse_drawing(child)
        elif tag == "textbox":
            # print("\n\n--\n")
            parse_tag(child, depth + 1)
            # print("\n\n--\n")
        elif tag == "lastRenderedPageBreak":
            print('\n\n<div class="break"></div>\n')
        else:
            # if tag not in IGNORE:
            #     print("*** unkown", f"<{tag}>")
            if tag in SHOW:
                print("#", tag)
            parse_tag(child, depth + 1)
            
def parse_drawing(root):
    tags = root.xpath(".//cNvPr")
    if len(tags) == 1 and "id" in tags[0].attrib:
        id = tags[0].attrib["id"]
        print(f'<img src="image{id}.png">')
    else:
        print(f"*** {len(tags)} pictures")

def get_tagname(s):
    index = s.find("}")
    if index != -1:
        return s[(index + 1):]
    return "?";

def read_file(file):
    with open(file, mode="rb") as f:
        return f.read()

def decode_shape(gfxdata, file):
    decoded = base64.b64decode(gfxdata.replace("&#xA;", "\n"))
    open(file, "wb").write(decoded)

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage: python docx2md <word.docx>")
        exit()
    file = sys.argv[1]
    if not os.path.exists(file) :
        print(file, "does not exit.")
        exit()
    parse_docx(file)


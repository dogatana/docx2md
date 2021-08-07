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
    save_xml(file, docx.document())
    doc = etree.fromstring(docx.document())
    strip_ns_prefix(doc)
    # resolver = NamespaceResolver(doc.nsmap)

    body = doc.xpath("//body")[0]
    parse_tag(body, 0)

def save_xml(file, text):
    # save for debug
    xml_file = os.path.splitext(file)[0] + ".xml"
    if os.path.exists(xml_file):
        return
    open(xml_file, "wb").write(text)
    print("save to", xml_file )

def strip_ns_prefix(tree):
    #xpath query for selecting all element nodes in namespace
    query = "descendant-or-self::*[namespace-uri()!='']"
    #for each element returned by the above xpath query...
    for element in tree.xpath(query):
        #replace element name with its local name
        element.tag = etree.QName(element).localname
    return tree

STOP_PARSING = [
    "pPr", "rPr", "sectPr", "pStyle", "wPr", "numPr", "ind", "shapetype", "tab"
]
NOT_PROCESS = [
    "shape", "txbxContent", "pict"
]
def parse_tag(root, depth):
    if root.tag in STOP_PARSING:
        return
    for child in root.getchildren():
        if  child is None:
            print("** none")
            continue
        tag = child.tag
        if tag in STOP_PARSING:
            continue
        if tag == "p":
            parse_p(child, depth + 1)
        elif tag == "r":
            parse_tag(child, depth + 1)
        elif tag == "br":
            if "page" in child.attrib.values():
                print('<div class="page"></div>')
            else:
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
            if tag not in NOT_PROCESS:
                print("#", tag) 
            parse_tag(child, depth + 1)
            
def parse_p(root, depth):
    """ parse paragraph """
    numPr = get_first_element(root, "./pPr/numPr")
    if numPr:
        # OL or UL
        ilvl = get_val(get_first_element(numPr, "./ilvl"))
        numId = get_val(get_first_element(numPr, "./numId")) # "1" for UL, "2" for OL
        print("    " * int(ilvl), end="")
        print("* " if numId == "1" else "1. ", end="")
    for child in root.getchildren():
        parse_tag(child, depth + 1)
    print("")

def get_first_element(tree, xpath):
    tags = tree.xpath(xpath)
    return tags[0] if len(tags) > 0 else None
    
def get_val(tag):
    for key, value in tag.attrib.items():
        if key.endswith("}val"):
            return value
    return None

def parse_drawing(root):
    tags = root.xpath(".//cNvPr")
    if tags and "id" in tags[0].attrib:
        id = tags[0].attrib["id"]
        print(f'<img src="image{id}.png">')
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
    if not os.path.exists(file) :
        print(file, "does not exit.")
        exit()
    parse_docx(file)


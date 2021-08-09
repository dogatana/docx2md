import sys
import os.path

from docxfile import DocxFile, DocxFileError
from converter import Converter


def parse_docx(file):
    docx = DocxFile(file)
    xml_text = docx.document()
    save_xml(file, xml_text)

    converter = Converter(xml_text)
    md_text = converter.convert()

    print("-" * 10, md_text, "-" * 10, sep="\n")
    open("out.md", "w", encoding="utf-8").write(md_text)


def save_xml(file, text):
    # save for debug
    xml_file = os.path.splitext(file)[0] + ".xml"
    if os.path.exists(xml_file):
        return
    open(xml_file, "wb").write(text)
    print("save to", xml_file)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage: python docx2md <word.docx>")
        exit()
    file = sys.argv[1]
    if not os.path.exists(file):
        print(file, "does not exit.")
        exit()
    parse_docx(file)

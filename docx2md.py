import sys
import os
import os.path

from docxfile import DocxFile, DocxFileError
from converter import Converter
from docxresources import DocxResources
from mediasaver import MediaSaver


def parse_docx(file):
    target_dir, _ = os.path.splitext(file)
    os.makedirs(target_dir, exist_ok=True)

    md_text = convert(file, target_dir, debug=True)
    save_md(os.path.join(target_dir, "README.md"), md_text)

def convert(file, target_dir, save_images=True, debug=True):
    docx = DocxFile(file)
    xml_text = docx.document()

    if debug:
        xml_file = os.path.splitext(os.path.basename(file))[0] + ".xml"
        save_xml(os.path.join(target_dir, xml_file), xml_text)

    rel_text = docx.read("word/_rels/document.xml.rels")
    res = DocxResources(rel_text)

    if save_images:
        saver = MediaSaver(docx, target_dir)
        saver.save()

    converter = Converter(xml_text, res)
    md_text = converter.convert()

    return md_text



def save_xml(file, text):
    xml_file = os.path.splitext(file)[0] + ".xml"
    if os.path.exists(xml_file):
        return
    open(xml_file, "wb").write(text)
    print("# save to", xml_file)

def save_md(file, text):
    open(file, "w", encoding="utf-8").write(text)
    print(f"# save {file}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage: python docx2md <word.docx>")
        exit()
    file = sys.argv[1]
    if not os.path.exists(file):
        print(file, "does not exit.")
        exit()
    parse_docx(file)

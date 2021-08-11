import sys
import os
import os.path
import argparse

from docxfile import DocxFile, DocxFileError
from converter import Converter
from docxmedia import DocxMedia

PROG = "docx2md"
VERSION = "1.0.0"


def main():
    args = parse_args()
    docx = create_docx(args.src)
    check_target_dir(args.dst)
    target_dir = os.path.dirname(args.dst)

    xml_text = docx.document()
    if args.debug:
        # save word/document.xml in docx
        xml_file = os.path.join(target_dir, "document.xml")
        print(f"# save {xml_file}")
        save_xml(xml_file, xml_text)
        with open(xml_file, mode="wb") as f:
            f.write(xml_text)

    media = DocxMedia(docx)
    md_text = convert(docx, target_dir, media, args.md_table)
    save_md(args.dst, md_text)
    media.save(target_dir)


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument("src", metavar="SRC.docx", help="Microsoft Word file to read")
    parser.add_argument(
        "dst",
        metavar="DST.md",
        help="Markdown file to write",
        default="./README.md",
    )
    parser.add_argument(
        "-m",
        "--md_table",
        action="store_true",
        help="use Markdown table notation instead of <table>",
        default=False,
    )
    parser.add_argument(
        "-v",
        "--version",
        help="show version",
        action="version",
        version=f"{PROG} {VERSION}",
    )
    parser.add_argument("--debug", help="for debug", action="store_true", default=False)
    return parser.parse_args()


def create_docx(file):
    try:
        return DocxFile(file)
    except Exception as e:
        print(e)
        exit()


def check_target_dir(file):
    dir = os.path.dirname(file)
    if dir == "":
        return
    if os.path.isdir(dir):
        return
    if os.path.exists(dir):
        print(f"cannot write to {file}")
        exit()
    os.makedirs(dir)


def convert(docx, target_dir, media, use_md_table):
    xml_text = docx.document()
    converter = Converter(xml_text, media, use_md_table)
    return converter.convert()


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
    main()

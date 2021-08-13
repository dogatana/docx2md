import sys
import os
import os.path
import argparse

from .docxfile import DocxFile, DocxFileError
from .converter import Converter
from .docxmedia import DocxMedia

PROG = "docx2md"
import docx2md

VERSION = f"{PROG} {docx2md.__version__}"


def main():
    args = parse_args()
    docx = create_docx(args.src)
    check_target_dir(args.dst)
    target_dir = os.path.dirname(args.dst)

    xml_text = docx.document()
    if args.debug:
        save_keyfile(docx, target_dir)

    media = DocxMedia(docx)
    md_text = convert(docx, target_dir, media, args.md_table)
    save_md(args.dst, md_text)

    media.save(target_dir)


def parse_args():
    parser = argparse.ArgumentParser(prog=PROG)
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
        version=VERSION,
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


def save_keyfile(docx, target_dir):
    def save_xml(file, text):
        print("# save", file)
        with open(file, "wb") as f:
            f.write(text)

    save_xml(os.path.join(target_dir, "document.xml"), docx.document())
    save_xml(os.path.join(target_dir, "document.xml.rels"), docx.rels())


def convert(docx, target_dir, media, use_md_table):
    xml_text = docx.document()
    converter = Converter(xml_text, media, use_md_table)
    return converter.convert()


def save_md(file, text):
    print("# save", file)
    with open(file, "w", encoding="utf-8") as f:
        f.write(text)


if __name__ == "__main__":
    main()

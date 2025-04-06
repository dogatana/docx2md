from .converter import Converter
from .docxfile import DocxFile
from .docxmedia import DocxMedia


def do_convert(docx_file: str, use_md_table=False) -> str:
    """convert docx_file to Markdown text and return it"""
    try:
        docx = DocxFile(docx_file)
        media = DocxMedia(docx)
        converter = Converter(docx.document(), media, use_md_table)
        return converter.convert()
    except Exception as e:
        return f"Exception: {e}"

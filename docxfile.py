import os
import os.path
import zipfile


class DocxFileError(Exception):
    pass


class DocxFile:
    def __init__(self, file):
        if not os.path.isfile(file):
            raise FileNotFoundError(f"{file} not found.")
        if not zipfile.is_zipfile(file):
            raise DocxFileError(f"{file} is not a zip file.")

        self.docx = zipfile.ZipFile(file)
        if "word/document.xml" not in self.namelist():
            raise DocxFileError(f"{file} is not a .docx file.")

    def document(self):
        if self.docx is None:
            raise FileNotFoundError
        return self.docx.read("word/document.xml")

    def rels(self):
        if self.docx is None:
            raise FileNotFoundError
        return self.docx.read("word/_rels/document.xml.rels")

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

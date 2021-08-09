import os
import os.path
import zipfile


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

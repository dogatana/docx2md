import os
import os.path
from PIL import Image


class MediaSaver:
    def __init__(self, docxfile, target_dir):
        self.docx = docxfile
        self.target_dir = os.path.join(target_dir, "media")

    def save(self):
        files = [f for f in self.docx.namelist() if f.lower().startswith("word/media/")]
        if len(files) == 0:
            return

        os.makedirs(self.target_dir, exist_ok=True)

        for file in files:
            self.save_image(file)

    def save_image(self, file):
        basename = os.path.basename(file)
        target_file = os.path.join(self.target_dir, basename)

        self.extract_file(file, target_file)
        self.correct_format(target_file)

    def extract_file(self, src, dst):
        print(f"# extract {dst}")
        with open(dst, "wb") as f:
            f.write(self.docx.read(src))

    def correct_format(self, file):
        base, ext = os.path.splitext(file)
        if ext.lower() in [".png", ".bmp", ".jpg", "jpeg"]:
            return

        print(f"# convert {file} to {base}.png")
        im = Image.open(file)
        im.save(base + ".png")
        os.unlink(file)

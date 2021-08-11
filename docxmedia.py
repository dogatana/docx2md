import os.path
import collections

from lxml import etree

import utils


class Media:
    def __init__(self, id, path):
        self.path = self.alt_path = path
        self.id = id
        self.use_alt = False
        self.check_alt(path)

    def check_alt(self, path):
        base, ext = os.path.splitext(path)
        if ext.lower() not in ".png .gif .bmp .jpg .jpeg".split():
            self.use_alt = True
            self.alt_path = base + ".png"

    def __repr__(self):
        return f"Media(id={self.id}, path={self.path}, alt_path={self.alt_path}, use_alt={self.use_alt}"


class DocxMedia:
    def __init__(self, docx):
        self.docx = docx
        self.hash = {}

        root = etree.fromstring(docx.rels())
            utils.strip_ns_prefix(root)
            self.parse_tree(root)

    def parse_tree(self, root):
        for tag in root.xpath("//Relationship"):
            path = tag.attrib.get("Target")
            if not path:
                continue
            if not path.startswith("media/"):
                continue
            id = tag.attrib.get("Id")
            if id:
                self.hash[id] = Media(id, path)

    def __contains__(self, id):
        return id in self.hash

    def __getitem__(self, id):
        return self.hash[id]

    def get(self, id):
        return self.hash.get(id)

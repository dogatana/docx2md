from lxml import etree


class DocxResources:
    def __init__(self, xml_text):
        self.hash = {}
        if xml_text is not None:
            root = etree.fromstring(xml_text)
            self.parse_tree(root)

    def parse_tree(self, root):
        for node in root.getchildren():
            id = ""
            path = ""
            for k, v in node.attrib.items():
                if k == "Id" or k.endswith("}Id"):
                    id = v
                if k == "Target" or k.endswith("}Target"):
                    path = v
            if id and path:
                self.hash[id] = path

    def get(self, id):
        return self.hash.get(id)

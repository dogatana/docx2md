import re

from lxml import etree


def strip_ns_prefix(tree):
    # xpath query for selecting all element nodes in namespace
    query = "descendant-or-self::*[namespace-uri()!='']"
    # for each element returned by the above xpath query...
    for element in tree.xpath(query):
        # replace element name with its local name
        element.tag = etree.QName(element).localname
        # remove prefix from attribute name
        stripped_attr = {
            re.sub(r"\{.*?\}", "", k): v for k, v in element.attrib.items()
        }
        element.attrib.clear()
        element.attrib.update(stripped_attr)

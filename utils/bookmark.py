from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree


class NAMESPACES:
    DOCX = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def add_bookmark(paragraph, bookmark_name):
    bookmark_start = OxmlElement('w:bookmarkStart')
    bookmark_start.set(qn('w:id'), str(id(paragraph)))
    bookmark_start.set(qn('w:name'), bookmark_name)

    bookmark_end = OxmlElement('w:bookmarkEnd')
    bookmark_end.set(qn('w:id'), str(id(paragraph)))

    paragraph._p.append(bookmark_start)
    paragraph._p.append(bookmark_end)


def create_bookmarks(name: str):
    bookmark_start = etree.Element("{" + NAMESPACES.DOCX['w'] + "}bookmarkStart", nsmap=NAMESPACES.DOCX)
    bookmark_start.attrib["{" + NAMESPACES.DOCX['w'] + '}id'] = str(hash(name))
    bookmark_start.attrib["{" + NAMESPACES.DOCX['w'] + '}name'] = name

    bookmark_end = etree.Element("{" + NAMESPACES.DOCX['w'] + "}bookmarkEnd", nsmap=NAMESPACES.DOCX)
    bookmark_end.attrib["{" + NAMESPACES.DOCX['w'] + '}id'] = str(hash(name))
    bookmark_end.attrib["{" + NAMESPACES.DOCX['w'] + '}name'] = name

    return bookmark_start, bookmark_end

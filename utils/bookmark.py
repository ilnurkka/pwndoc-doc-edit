from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def add_bookmark(paragraph, bookmark_name):
    bookmark_start = OxmlElement('w:bookmarkStart')
    bookmark_start.set(qn('w:id'), str(id(paragraph)))
    bookmark_start.set(qn('w:name'), bookmark_name)

    bookmark_end = OxmlElement('w:bookmarkEnd')
    bookmark_end.set(qn('w:id'), str(id(paragraph)))

    paragraph._p.append(bookmark_start)
    paragraph._p.append(bookmark_end)


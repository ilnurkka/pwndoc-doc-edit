from lxml import etree

# <w:highlight w:val="yellow"/>


class NAMESPACES:
	DOCX = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def highlight_clear(document):
	for h in document._element.findall('.//w:highlight', NAMESPACES.DOCX):
		h.getparent().remove(h)
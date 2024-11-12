import io
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches, Cm

from utils import find_images_and_captions, replace_image_references


def document_preparing(document: Document):
	image_counter = 1
	in_section_6 = False
	need_to_delete_next = False
	next_p_add_art = False
	list_paragraph_to_check = ["Описание:", "Наблюдения:", "Ссылки:", "Доказательства:", "Затронутые активы:",
							   "Рекомендации:"]

	for paragraph in document.paragraphs:
		if paragraph.style.name == 'Normal':
			paragraph.style = document.styles['Основной текст icl']

		if paragraph.text.find("оценивается как") != -1:
			start = paragraph.text.find("оценивается как")
			picture = paragraph.text[start + 16:].split(",")[0]

		if next_p_add_art:
			run = paragraph.insert_paragraph_before()
			run.alignment = 1
			run.add_run().add_picture(f"arts/{picture}.png", width=Inches(6))
			next_p_add_art = False

		if paragraph.text.find("От внешнего нарушителя ресурса") != -1:
			next_p_add_art = True

		if "Детальное описание хода работ и результатов" in paragraph.text and paragraph.style.name == "Heading 1":
			in_section_6 = True

		if in_section_6:
			if paragraph.text == "" and reserv_paragraph.text in list_paragraph_to_check:
				p = reserv_paragraph._element
				p.getparent().remove(p)
				p._element = None
				p_ = paragraph._element
				p_.getparent().remove(p_)
				p_._element = None

			if need_to_delete_next:
				p = paragraph._element
				p.getparent().remove(p)
				p._element = None
				need_to_delete_next = False

			if paragraph.text == "CVSS балл:  ()":
				p = paragraph._element
				p.getparent().remove(p)
				p._element = None
				need_to_delete_next = True

			if paragraph.style.name == 'Caption':
				paragraph.text = f"Рисунок 6.{image_counter} - {paragraph.text}"
				image_counter += 1

			reserv_paragraph = paragraph

	images = find_images_and_captions(document)
	replace_image_references(document, images)

	# Работа с определениями
	table = document.tables[0]

	for i, column in enumerate(table.columns):
		if i == 0:
			column.width = Cm(2)  # Первый столбец - 2 см
		elif i == 1:
			column.width = Cm(0.5)  # Второй столбец - 0.5 см
		elif i == 2:
			column.width = Cm(13)  # Третий столбец - 13 см

	for row in table.rows:
		row.cells[0].width = Cm(2)
		row.cells[1].width = Cm(0.5)
		row.cells[2].width = Cm(13)

	table.style = 'pwndoc-table'

	for row in table.rows:
		for cell in row.cells:
			cell._element.get_or_add_tcPr().append(parse_xml(
				r'<w:tcBorders {}><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'.format(
					nsdecls('w'))))

	# Работа с терминами
	table2 = document.tables[1]

	for i, column in enumerate(table2.columns):
		if i == 0:
			column.width = Cm(4)  # Первый столбец - 2 см
		elif i == 1:
			column.width = Cm(0.5)  # Второй столбец - 0.5 см
		elif i == 2:
			column.width = Cm(11)  # Третий столбец - 13 см

	for row in table2.rows:
		row.cells[0].width = Cm(4)
		row.cells[1].width = Cm(0.5)
		row.cells[2].width = Cm(11)

	table2.style = 'pwndoc-table'

	for row in table2.rows:
		for cell in row.cells:
			cell._element.get_or_add_tcPr().append(parse_xml(
				r'<w:tcBorders {}><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>'.format(
					nsdecls('w'))))


d
import html

from lxml import etree

from .bookmark import add_bookmark, create_bookmarks
from .hyperlink import add_hyperlink, create_hyperlink


class NAMESPACES:
    DOCX = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def replace_image_references_old(document, image_map):
    for i, paragraph in enumerate(document.paragraphs):
        last_paragraph = None
        for key, value in image_map.items():
            search_text = f'[см. {key}]'
            replacement_text = f'{value}'

            if search_text in paragraph.text:
                parts = paragraph.text.split(search_text)

                p = paragraph._p
                if p is not None and p.getparent() is not None:
                    p.getparent().remove(p)

                next_paragraph = document.paragraphs[i] if i < len(document.paragraphs) - 1 else None

                if next_paragraph is not None:
                    new_paragraph = next_paragraph.insert_paragraph_before()
                else:
                    new_paragraph = document.add_paragraph()  # Вставка в конец, если последний абзац

                add_bookmark(new_paragraph, key)

                new_paragraph.add_run(parts[0])

                add_hyperlink(new_paragraph, key, replacement_text)

                if len(parts) > 1:
                    new_paragraph.add_run(parts[1])

                new_paragraph.style = document.styles['Основной текст icl']


def replace_image_references(document, images):
    for i, p in enumerate(document._element.findall('.//w:p', NAMESPACES.DOCX)):

        def clear_searches(replaces, text):
            new_text = text
            for search, replace, key in replaces:
                new_text = new_text.replace(search, '')
            return new_text

        replaces = []
        for key, value in images.items():
            search_text = f'[см. {key}]'
            replacement_text = f'{value}'

            if p.text is not None and search_text in p.text:
                replaces.append((search_text, replacement_text, key))

        if len(replaces) > 0:
            for _run in p.findall('.//w:r', NAMESPACES.DOCX):
                new_runs = []
                for search, replace, key in replaces:
                    for run in [_run] + new_runs:
                        text = run.text if run.text is not None else ''.join(run.xpath('.//text()'))
                        if search in text:
                            parts = text.split(search)

                            if run.text is not None:
                                run.text = parts[0]
                            else:
                                t = run.find(".//w:t", NAMESPACES.DOCX)
                                if t is not None and t.text is not None:
                                    t.text = parts[0]

                            hyperlink = create_hyperlink(key, replace)
                            run.addnext(hyperlink)

                            new_run = etree.Element("{" + NAMESPACES.DOCX['w'] + "}r", nsmap=NAMESPACES.DOCX)

                            t = etree.SubElement(new_run, "{" + NAMESPACES.DOCX['w'] + "}t", nsmap=NAMESPACES.DOCX)
                            t.text = parts[1]
                            # new_run.text = parts[1]

                            hyperlink.addnext(new_run)
                            new_runs.append(new_run)
                            break


def find_images_and_captions(document):
    images = {}

    for paragraph in document.paragraphs:
        if 'Рисунок' in paragraph.text:
            caption = paragraph.text.split('-')[0].strip()
            name = paragraph.text.split('-')[1].strip()
            images[name] = caption

    return images

from .bookmark import add_bookmark
from .hyperlink import add_hyperlink


def replace_image_references(document, image_map):
    for i, paragraph in enumerate(document.paragraphs):
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


def find_images_and_captions(document):
    images = {}

    for paragraph in document.paragraphs:
        if 'Рисунок' in paragraph.text:
            caption = paragraph.text.split('-')[0].strip()
            name = paragraph.text.split('-')[1].strip()
            images[name] = caption

    return images

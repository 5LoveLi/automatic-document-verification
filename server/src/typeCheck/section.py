from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

import errorList

# Проверка полей
def indent(doc):
    for i, section in enumerate(doc.sections):

        if section.top_margin.inches != 0.7875:
            s = i + 1
            # print(errorList.Error.check(float(section.top_margin.inches) * 2.54, str(s), 12))
            return False

        if section.left_margin.inches != 1.18125:
            s = i + 1
            # print(errorList.Error.check(float(section.left_margin.inches) * 2.54, str(s), 12))
            return False

        if section.right_margin.inches != 0.5909722222222222:
            s = i + 1
            # print(errorList.Error.check(float(section.right_margin.inches) * 2.54,str(s), 12))
            return False

def structure(doc):
    l = ['F', 'F', 'F', 'F', 'F', 'F']
    expected_texts = ["СОДЕРЖАНИЕ", "РЕФЕРАТ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "БИБЛИОГРАФИЧЕСКИЙ ИСТОЧНИК", "ПРИЛОЖЕНИЯ"]

    for paragraph in doc.paragraphs:
        alignment = paragraph.alignment
        if paragraph.style.name == 'Heading 1' and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            t = paragraph.text
            if t in expected_texts:
                index = expected_texts.index(t)
                l[index] = 'T'

    if l[0] == 'F':
        print("\033[31m {}\033[0m".format("Содержание отсутствует"))
    if l[1] == 'F':
        print("\033[31m {}\033[0m".format("Реферат отсутствует"))
    if l[2] == 'F':
        print("\033[31m {}\033[0m".format("Введение отсутствует"))
    if l[3] == 'F':
        print("\033[31m {}\033[0m".format("Заключение отсутствует"))
    if l[4] == 'F':
        print("\033[31m {}\033[0m".format("Библиографический источник отсутствует"))
    if l[5] == 'F':
        print("\033[31m {}\033[0m".format("Приложения отсутствуют"))



def numbering(document):
    total_pages = 0  # Общее количество страниц

    # Нумерация страниц по всему тексту
    for section in document.sections:
        if not section.footer.is_linked_to_previous:
            total_pages += section.page_count

    # Нумерация иллюстраций и таблиц
    elements = list(document.inline_shapes) + document.tables
    total_pages += len([element for element in elements if element.width > 595.3 or element.height > 841.9])

    # Проверка соответствия требованиям
    for paragraph in document.paragraphs:
        if paragraph.text.strip():  # Игнорирование пустых абзацев
            run = paragraph.runs[0]
            if run.bold:  # Проверка титульного листа
                continue
            if run.text.isnumeric():  # Проверка нумерации арабскими цифрами
                if int(run.text) != total_pages:
                    return False
                total_pages += 1

    return True
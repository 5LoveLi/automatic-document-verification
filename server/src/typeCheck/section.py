from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

def structure(doc):

    error = {
    "РЕФЕРАТ" : 'Ошибка структуры: Реферат отсутствует.',
    "СОДЕРЖАНИЕ" : 'Ошибка структуры: Содержание отсутствует.',
    "ВВЕДЕНИЕ" : 'Ошибка структуры: Введение отсутствует.',
    "ЗАКЛЮЧЕНИЕ" : 'Ошибка структуры: Заключение отсутствует.',
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" : 'Ошибка структуры: Список использованных источников отсутствует.',
    "ПРИЛОЖЕНИЯ" : 'Ошибка структуры: Приложения отсутствуют.'
    }

    error_list = []
    expected_texts = [ "РЕФЕРАТ", "СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "ПРИЛОЖЕНИЯ"]

    for paragraph in doc.paragraphs:
        if paragraph.style.name == 'Heading 1' and paragraph.alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
            t = paragraph.text
            if t in expected_texts:
                expected_texts.remove(t)

    for i in expected_texts:
        error_list.append(error[i])

    return error_list



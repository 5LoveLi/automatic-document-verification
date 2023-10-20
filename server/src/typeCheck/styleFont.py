import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import RGBColor

from src import errorList



def font(paragraph):
    error_list = []
    if paragraph.style.name != 'Heading 1' and paragraph.style.name != 'Heading 2' and paragraph.style.name != 'Heading 3':
        for run in paragraph.runs:
            
            font_name = run.font.name
            if  font_name != 'Times New Roman':
                error_list.append(errorList.error[13])
                print("\033[31m {}\033[0m" .format(errorList.error[13]))

            font_size = run.font.size
            if font_size != 14:
                error_list.append(errorList.error[14])
                print("\033[31m {}\033[0m" .format(errorList.error[14]))

            font_color = run.font.color.rgb
            if font_color is not None and font_color != (0, 0, 0):
                error_list.append(errorList.error[16])
                print("\033[31m {}\033[0m" .format(errorList.error[16]))

            if not(not run.font.italic and not run.font.bold and not run.font.underline):
                error_list.append(errorList.error[17])
                print("\033[31m {}\033[0m" .format(errorList.error[17]))

        if paragraph.paragraph_format.line_spacing != 1.5:
            p = paragraph.paragraph_format.line_spacing
            error_list.append(errorList.error[15])
            print("\033[31m {}\033[0m" .format(errorList.error[15]))
    
    return error_list


def all_header(paragraph) -> list: 
    error_list = []
    section_text = paragraph.text.strip()

    if section_text.endswith('.'):
                error_list.append(errorList.error[18])
                print("\033[31m {}\033[0m" .format(errorList.error[18]))

        # Проверка строчных букв
    if not any(letter.islower() for letter in section_text):
        error_list.append(errorList.error[19])
        print("\033[31m {}\033[0m" .format(errorList.error[19]))

    if section_text[0].islower():
        error_list.append(errorList.error[20])
        print("\033[31m {}\033[0m" .format(errorList.error[20]))

    # Проверка полужирного шрифта
    if not paragraph.runs[0].bold:
        error_list.append(errorList.error[21])
        print("\033[31m {}\033[0m" .format(errorList.error[21]))

    # Проверка расположения слева
    if paragraph.alignment != None:
        error_list.append(errorList.error[22])
        print("\033[31m {}\033[0m" .format(errorList.error[22]))

    # Проверка отсутствия подчеркивания
    if paragraph.runs[0].underline:
        error_list.append(errorList.error[23])
        print("\033[31m {}\033[0m" .format(errorList.error[23]))

    # Проверка цвета заголовка
    if paragraph.runs[0].font.color.rgb != RGBColor(0, 0, 0):
        error_list.append(errorList.error[16])
        print("\033[31m {}\033[0m" .format(errorList.error[16]))

    # Проверка абзацного отступа
    if paragraph.paragraph_format.first_line_indent != 450215:
        error_list.append(errorList.error[24])
        print("\033[31m {}\033[0m" .format(errorList.error[24]))
        # print("заголовков разделов")

    if paragraph.paragraph_format.line_spacing != 1.5:
        p = paragraph.paragraph_format.line_spacing
        error_list.append(errorList.error[15])
        print("\033[31m {}\033[0m" .format(errorList.error[15]))

    for run in paragraph.runs:
        font_name = run.font.name
        if font_name != 'Times New Roman':
            error_list.append(errorList.error[15])
            print("\033[31m {}\033[0m" .format(errorList.error[15]))

    return error_list

    

def header(paragraph):
    
    # section_num = 1
    error_list = []
    alignment = paragraph.alignment

    # Проверка заголовков разделов
    if paragraph.style.name == 'Heading 1' and alignment != WD_PARAGRAPH_ALIGNMENT.CENTER:
        error_list += all_header(paragraph)
        print('Проверены 1-ый заголовок')

    # Проверка заголовков подразделов
    if paragraph.style.name == 'Heading 2':
        error_list += all_header(paragraph)
        print('Проверен 2-ой заголовок')

    if paragraph.style.name == 'Heading 1' and alignment == WD_PARAGRAPH_ALIGNMENT.CENTER:
        section_text = paragraph.text.strip()

        # Проверка отсутствия точки в конце
        if section_text.endswith('.'):
            error_list.append(errorList.error[18])
            print(errorList.error[18])

        # Проверка строчных букв
        if not section_text.isupper():
            error_list.append(errorList.error[25])
            print(errorList.error[25])

        # Проверка полужирного шрифта
        if not paragraph.runs[0].bold:
            
            print(errorList.error[21])

        # Проверка отсутствия подчеркивания
        if paragraph.runs[0].underline:
            error_list.append(errorList.error[23])
            print(errorList.error[23])

        # Проверка цвета заголовка
        if paragraph.runs[0].font.color.rgb != RGBColor(0, 0, 0):
            error_list.append(errorList.error[16])
            print(errorList.error[16])

        for run in paragraph.runs:
            font_name = run.font.name
            if font_name != 'Times New Roman':
                error_list.append(errorList.error[13])
                print(errorList.error[13])

    
    return error_list

    # Проверка нумерации
    # if section_text[0] == str(section_num) + " ":
    #     print('Проверка нумерации')
    # print(section_text)
    # section_num += 1

    
def find_lists(doc):
    # print(errorList.error[13])
    for paragraph in doc.paragraphs:
        if (
            paragraph.style.name.startswith('List')
            and paragraph.style.type == 1
            and paragraph.style.paragraph_format.left_indent is not None
            and paragraph.style.paragraph_format.left_indent > 0
        ):
            print(paragraph.text)

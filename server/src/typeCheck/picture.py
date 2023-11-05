# from docx import Document
# from docx.oxml import ns
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH

from src import errorList
from src.typeCheck import styleFont



def extract_images_from_docx(document):
    num = 0
    
    error_list = []
    extract_text = False
    next_paragraph = False
    for paragraph in document.paragraphs:
        section_text = paragraph.text.strip()
        for run in paragraph.runs:
            if run._r.xml:
                xml = run._r.xml
                if '<wp:inline' in xml:
                    num += 1
                    print(f'картинка {num}:')
                    next_paragraph = True

        if next_paragraph:
            next_paragraph = False
            extract_text = True

        elif extract_text:
            text = paragraph.text
            if 'Рисунок' in text:
                extract_text = False

                if  re.match(r"Рисунок \d+\.\d+", text):
                    print('сверяем первую цифру с цифрой раздела ')

                elif  re.match(r"Рисунок \d+", text) is None:
                    error_list.append(errorList.error[43])
                    print(errorList.error[43])

                else:
                    number = re.findall(r'\d+', text)[0]
                    if int(number) != num:
                        error_list.append(errorList.error[44])
                        print(errorList.error[44])

                # Проверка стиля текста
                # error_list.append(styleFont.font(paragraph))

                # Проверка что описание по центру
                if paragraph.alignment != WD_ALIGN_PARAGRAPH.CENTER:
                    error_list.append(errorList.error[41])
                    print(errorList.error[41])
                
                # Проверка на точку в конце
                if section_text.endswith('.'):
                    error_list.append(errorList.error[42])
                    print(errorList.error[42])

                
            else:
                print(errorList.error[40])
                error_list.append(errorList.error[40])
                extract_text = False
















# def find_captions(doc):
#     # Получаем элемент body
#     body = doc.element.body
#
#     # Определяем пространство имен для элементов рисунков
#     nsdecls = {
#         'w': ns.nsmap['w'],
#         'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
#         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
#     }
#
#     # Перебираем все параграфы внутри body
#     for paragraph in body.iterchildren():
#         # Проверяем, содержит ли параграф элемент рисунка
#         drawing = paragraph.find('.//w:drawing', namespaces=nsdecls)
#         if drawing is not None:
#             # Находим элемент подписи рисунка
#             caption = drawing.xpath('.//pic:cNvPr', namespaces=nsdecls)
#             if caption:
#                 caption_name = caption[0].get('name')
#                 print(caption_name)

def explore_xml_structure(doc):
    # Получаем XML-представление документа
    xml = doc._element.xml

    # Выводим XML-представление документа
    print(xml)
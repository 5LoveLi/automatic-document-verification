from docx import Document
from docx.oxml import ns

def extract_images_from_docx(document):
    extract_text = False
    next_paragraph = False
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if run._r.xml:
                xml = run._r.xml
                if '<wp:inline' in xml:
                    extract_text = True
                    next_paragraph = True
        if next_paragraph:
            next_paragraph = False
            extract_text = True
        elif extract_text:
            text = paragraph.text
            if text != '':
                # print("Изображение подписано: ",text)
                extract_text = False
            else:
                print("Изображение не подписано!")
                break
















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
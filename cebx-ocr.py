# encoding=utf-8
import re
import easyocr
import docx
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from wand.image import Image
import os


CEBX_DIR = 'D:/Data/code/policy-analysis/input-files/cebx/'
CEBX_PDF_FILE_SUFFIX = '.pdf'
CEBX_TEST_NAME = 'P020180709553071286736'


def pdf_to_pic(file_path):
    image_pdf = Image(filename=file_path, resolution=300)
    image_jpg = image_pdf.convert('jpeg')
    imgs = []
    for img in image_jpg.sequence:
        img_page = Image(image=img)
        imgs.append(img_page.make_blob('jpeg'))
    return imgs


def pic_ocr_to_text(img_list):
    reader = easyocr.Reader(['ch_sim', 'en'])
    text_list = []
    for img in img_list:
        result = reader.readtext(img)
        text = ''
        for item in result:
            text += item[1]
        text_list.append(text)
    return text_list


def text_to_docx(dir, name, text_list):
    file = docx.Document()
    file.styles['Normal'].font.name = u'宋体'
    file.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    file.styles['Normal'].font.size = Pt(10.5)
    file.styles['Normal'].font.color.rgb = RGBColor(0, 0, 0)

    for text in text_list:
        file.add_paragraph(text)

    file.save(dir + name + '.docx')


def pdf_to_docx(dir, suffix):
    files = os.listdir(dir)
    for file in files:
        if os.path.isdir(file):
            continue
        if file.find('~$') != -1:
            continue
        if os.path.splitext(file)[1] != suffix:
            continue
        print('------start convert pdf:' + file)
        text_list = pic_ocr_to_text(pdf_to_pic(
            dir+file))
        text_to_docx(dir, os.path.splitext(file)[0], text_list)


if __name__ == '__main__':
    pdf_to_docx(CEBX_DIR, CEBX_PDF_FILE_SUFFIX)

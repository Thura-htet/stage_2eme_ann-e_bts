import pandas as pd
from itertools import product
import os

from PIL import Image

from barcode import Code128
from barcode import EAN13
from barcode.writer import ImageWriter
from barcode.writer import SVGWriter

from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph

from config import *


# TODO : convert into helper class
configs = read_config("config.json")  #  move this to main and call ti


def get_page_size(page_size):
    if page_size.lower() == "a4":
        return A4
    elif page_size.lower() == "a3":
        return A3
    elif page_size.lower() == "letter":
        return letter
    elif page_size.lower() == "legal":
        return legal
    else:
        return A4


DOC_WIDTH, DOC_HEIGHT = get_page_size(configs['page_size'])

top, left = configs['margin']['top'] * mm, configs['margin']['left'] * mm
hor_space, ver_space = DOC_WIDTH - left * 2, DOC_HEIGHT - top * 2
rows, cols = configs['dimensions']['rows'], configs['dimensions']['cols']
card_width, card_height = hor_space / cols, ver_space / rows


def read_file(file, sheet_name=None):
    file_ext = file.split(".")[-1]
    if file_ext == "csv":
        with open(file, newline='') as csvfile:
            df = pd.read_csv(csvfile, sep='\t')
            stock = df.to_dict('records')
            return stock

    elif file_ext == "xlsm" or file_ext == "xlsx":
        try:
            with pd.ExcelFile(file) as xls:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=0)
                stock = df.to_dict('records')
                return stock
        except Exception as e:
            print(f"Error reading {file}: {e}")


def remove_barcode_padding(barcode_file):
    image = Image.open(barcode_file)

    bbox = image.getbbox()
    cropped_image = image.crop(bbox)

    cropped_image.save(barcode_file)


def generate_product_label(pdf_canvas, product, x, y):
    barcode_file = f"barcode_{str(product['Référence']).strip().replace('/', '_')}.png"
    with open(barcode_file, "wb") as outfile:
        Code128(str(product['Référence']), writer=ImageWriter()).write(outfile)

    rayon_paragraph_style = ParagraphStyle(
        'CustomStyle',
        parent=getSampleStyleSheet()['Normal'],
        fontName='Helvetica',
        fontSize=10
    )
    rayon_paragraph = Paragraph(f"{product['Code Rayon']}", rayon_paragraph_style)
    rayon_paragraph.wrapOn(pdf_canvas, card_width * 0.3, card_height)

    price_paragraph_style = ParagraphStyle(
        'CustomStyle',
        parent=getSampleStyleSheet()['Normal'],
        alignment=1,  # centre
        fontName='Helvetica-Bold',
        fontSize=20
    )
    price_paragraph = Paragraph(
        f"{product['PV MB']}" if 'PV MB' in product else f"{product['PV MB CALCULE COEFF']} CHF",
        price_paragraph_style
    )
    price_paragraph.wrapOn(pdf_canvas, card_width, card_height)

    desc_paragraph_style = ParagraphStyle(
        'CustomStyle',
        parent=getSampleStyleSheet()['Normal'],
        backColor=colors.white,
        fontName='Helvetica-Bold'
    )
    desc_paragraph = Paragraph(product['Désignation'], desc_paragraph_style)
    desc_paragraph.wrapOn(pdf_canvas, card_width, card_height)

    pdf_canvas.drawInlineImage(
        barcode_file,
        x + (card_width * 0.4), y,
        width=card_width * 0.6,
        height=card_height * 0.8
    )
    rayon_paragraph.drawOn(pdf_canvas, x, y + rayon_paragraph.height)
    price_paragraph.drawOn(pdf_canvas, x, y + card_height - price_paragraph.height)
    desc_paragraph.drawOn(pdf_canvas, x, (y + card_height - desc_paragraph.height - price_paragraph.height * 2))

    os.remove(barcode_file)


def generate_pdf(products, pdf_file, pagesize):
    sorted_products = sorted(products, key=lambda item: item["Code Rayon"])
    pdf_canvas = canvas.Canvas(pdf_file, pagesize=get_page_size(pagesize))
    rows = 7
    cols = 3
    y = 0

    for i, product in enumerate(sorted_products):
        if (i != 0) and (i % (rows * cols) == 0):
            pdf_canvas.showPage()

        row = (i // cols) % rows
        col = i % cols
        x = left + (col * card_width)
        # the +1 is because we are starting from the top of the document and the file is being written from the bottom.
        # we need to save the space for the height of one row
        y = DOC_HEIGHT - ((row + 1) * card_height) - top

        print(product)
        print()
        generate_product_label(pdf_canvas, product, x, y)

    pdf_canvas.save()

from io import BytesIO
import csv
from itertools import product
import os

from PIL import Image

from barcode import Code128
from barcode import EAN13
from barcode.writer import ImageWriter
from barcode.writer import SVGWriter

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph

A4_WIDTH, A4_HEIGHT = A4
LETTER_WIDTH, LETTER_HEIGHT = letter
left, right, top, bottom = 14 * mm, 14 * mm, 11 * mm, 11 * mm
hor_space, ver_space = A4_WIDTH - left - right, A4_HEIGHT - top - bottom
rows, cols = 7, 3
card_width, card_height = hor_space / cols, ver_space / rows


def read_csvfile(file):
    Stock = []
    with open(file, newline='') as csvfile:
        reader = csv.DictReader(csvfile, delimiter='\t')
        for row in reader:
            Stock.append(row)
    return Stock


def remove_barcode_padding(barcode_file):
    image = Image.open(barcode_file)

    bbox = image.getbbox()
    cropped_image = image.crop(bbox)

    cropped_image.save(barcode_file)


def generate_product_label(pdf_canvas, product, x, y):
    barcode_file = f"barcode_{product['Référence'].strip().replace('/', '_')}.png"
    with open(barcode_file, "wb") as outfile:
        Code128(product['Référence'], writer=ImageWriter()).write(outfile)

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
    price_paragraph = Paragraph(f"{product['PV MB']}", price_paragraph_style)
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
        x + (card_width * 0.5), y,
        width=card_width * 0.5,
        height=card_height * 0.5
    )

    rayon_paragraph.drawOn(pdf_canvas, x, y + rayon_paragraph.height)

    price_paragraph.drawOn(pdf_canvas, x, y + card_height - price_paragraph.height)

    desc_paragraph.drawOn(pdf_canvas, x, (y + card_height - desc_paragraph.height - price_paragraph.height*2))

    os.remove(barcode_file)

# TODO: factorize code
# the barcode png files come out in the same size
# the position for item info is all mixed up
# the extra space at the bottom of the pages are due to vertical margins


def generate_pdf(products, pdf_file, pagesize):
    print(products)
    sorted_products = sorted(products, key=lambda item: item["Code Rayon"])
    pdf_canvas = canvas.Canvas(pdf_file, pagesize=pagesize)
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
        y = A4_HEIGHT - ((row+1) * card_height) - bottom

        print(product)
        generate_product_label(pdf_canvas, product, x, y)

    pdf_canvas.save()


if __name__ == "__main__":
    Stock = read_csvfile('grille_stoko.csv')
    test_stock = [Stock[1]]
    generate_pdf(Stock, "test_doc.pdf", A4)

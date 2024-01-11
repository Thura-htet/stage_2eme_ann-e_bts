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
left, right, top, bottom = 7 * mm, 7 * mm, 14 * mm, 14 * mm
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
        x + (card_width * 0.3), y + desc_paragraph.height,
        width=card_width * 0.5,
        height=card_height * 0.5
    )

    desc_paragraph.drawOn(pdf_canvas, x, (y - desc_paragraph.height))

    price_paragraph_style = ParagraphStyle(
        'CustomStyle',
        parent=getSampleStyleSheet()['Normal'],
        alignment=1,  # centre
        fontName='Helvetica-Bold',
        fontSize=20
    )
    price_paragraph = Paragraph(f"{product['PV MB']}", price_paragraph_style)
    price_paragraph.wrapOn(pdf_canvas, card_width, card_height)
    price_paragraph.drawOn(pdf_canvas, x, y + price_paragraph.height + (3 * mm))

    rayon_paragraph_style = ParagraphStyle(
        'CustomStyle',
        parent=getSampleStyleSheet()['Normal'],
        fontName='Helvetica',
        fontSize=12
    )
    rayon_paragraph = Paragraph(f"{product['Code Rayon']}", rayon_paragraph_style)
    rayon_paragraph.wrapOn(pdf_canvas, card_width * 0.3, card_height)
    rayon_paragraph.drawOn(pdf_canvas, x, y + price_paragraph.height + (7 * mm) + rayon_paragraph.height)

    os.remove(barcode_file)

# TODO: fix the font sizes
# TODO: add margin to each page and centralize the cards
# TODO: factorize code
# the barcode png files come out in the same size
# the position for item info is all mixed up
# the extra space at the bottom of the pages are due to vertical margins


def generate_pdf(products, pdf_file, pagesize):
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
        print(f'row: {row},y: {y}')
        print()
        pdf_canvas.setStrokeColorRGB(0, 0, 0)  # Set line color to black
        pdf_canvas.setLineWidth(5)  # Set line width
        pdf_canvas.line(0, y, A4_WIDTH, y)  # Adjust y coordinate as needed

        generate_product_label(pdf_canvas, product, x, y)

    pdf_canvas.save()


if __name__ == "__main__":
    Stock = read_csvfile('grille_stoko.csv')
    generate_pdf(Stock, "test_doc.pdf", A4)

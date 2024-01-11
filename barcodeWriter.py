from io import BytesIO
import csv
from itertools import product
import os

from barcode import Code128
from barcode import EAN13
from barcode.writer import ImageWriter
from barcode.writer import SVGWriter

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib import colors
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph

A4_WIDTH, A4_HEIGHT = A4
LETTER_WIDTH, LETTER_HEIGHT = letter
rows, cols = 7, 3
card_width, card_height = A4_WIDTH / cols, A4_HEIGHT / rows

Stock = []
with open('grille_stoko.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile, delimiter='\t')
    for row in reader:
        Stock.append(row)

first_item = Stock[1]


def generate_product_label(pdf_canvas, product, x, y):
    print(product['Référence'])
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
        width=card_width * 0.7,
        height=card_height * 0.7
    )
    """
    desc_paragraph_style = ParagraphStyle(
        'CustomStyle',
        parent=getSampleStyleSheet()['Normal'],
        backColor=colors.white,
        fontName='Helvetica-Bold'
    )
    desc_paragraph = Paragraph(product['Désignation'], desc_paragraph_style)
    desc_paragraph.wrapOn(pdf_canvas, card_width, card_height)
    desc_paragraph.drawOn(pdf_canvas, 0, (card_height-desc_paragraph.height))
    """
    desc_paragraph.drawOn(pdf_canvas, x, (y - desc_paragraph.height))

    price_paragraph_style = ParagraphStyle(
        'CustomStyle',
        parent=getSampleStyleSheet()['Normal'],
        alignment=1,  # centre
        fontName='Helvetica-Bold',
        fontSize=24
    )
    price_paragraph = Paragraph(f"{product['PV MB']}", price_paragraph_style)
    price_paragraph.wrapOn(pdf_canvas, card_width, card_height)
    price_paragraph.drawOn(pdf_canvas, x, y + price_paragraph.height + (3 * mm))

    rayon_paragraph_style = ParagraphStyle(
        'CustomStyle',
        parent=getSampleStyleSheet()['Normal'],
        fontName='Helvetica',
        fontSize=16
    )
    rayon_paragraph = Paragraph(f"{product['Code Rayon']}", rayon_paragraph_style)
    rayon_paragraph.wrapOn(pdf_canvas, card_width * 0.3, card_height)
    rayon_paragraph.drawOn(pdf_canvas, x, y + price_paragraph.height + (7 * mm) + rayon_paragraph.height)

    os.remove(barcode_file)


def generate_pdf(products, pdf_file, pagesize):
    sorted_products = sorted(products, key= lambda item: item["Code Rayon"])
    pdf_canvas = canvas.Canvas(pdf_file, pagesize=pagesize)
    rows = 7
    cols = 3

    for i, product in enumerate(sorted_products):
        if not i or not (i % (rows * cols)):
            pdf_canvas.showPage()

        row = (i // cols) % rows
        col = i % cols
        x = col * card_width
        y = A4_HEIGHT - (row+1) * card_height

        generate_product_label(pdf_canvas, product, x, y)

    pdf_canvas.save()


pdf_canvas = canvas.Canvas("test.pdf", pagesize=A4)
generate_product_label(pdf_canvas, first_item, 0, 0)

pdf_canvas.showPage()
pdf_canvas.save()

generate_pdf(Stock, "test_doc.pdf", A4)

from io import BytesIO
import csv

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

Stock = []
with open('grille_stoko.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile, delimiter='\t')
    for row in reader:
        Stock.append(row)

first_item = Stock[1]
print(first_item)

with open(f"barcode_{first_item['Référence']}.png", "wb") as outfile:
    Code128(first_item['Référence'], writer=ImageWriter()).write(outfile)

pdf_canvas = canvas.Canvas("test.pdf", pagesize=A4)

card_width, card_height = A4_WIDTH / 3, A4_HEIGHT / 7

"""
first_text_object = pdf_canvas.beginText(0, (1*cm))
first_text_object.textLine("this is the first line.")

second_text_object = pdf_canvas.beginText(0, A4_HEIGHT-card_height)
second_text_object.textLine("this is the second line.")

pdf_canvas.drawInlineImage(f"barcode_{first_item['Référence']}.png", 0, A4_HEIGHT-card_height, width = A4_WIDTH / 3, height = A4_HEIGHT/ 7)
pdf_canvas.drawText(first_text_object)
pdf_canvas.drawText(second_text_object)
"""

desc_paragraph_style = ParagraphStyle(
    'CustomStyle',
    parent=getSampleStyleSheet()['Normal'],
    backColor=colors.white,
    fontName='Helvetica-Bold'
)
desc_paragraph = Paragraph(first_item['Désignation'], desc_paragraph_style)
desc_paragraph.wrapOn(pdf_canvas, card_width, card_height)

pdf_canvas.drawInlineImage(
    f"barcode_{first_item['Référence']}.png",
    (A4_WIDTH/3)*0.3, desc_paragraph.height,
    width=(A4_WIDTH/3)*0.7,
    height=(A4_HEIGHT/7)*0.7
)
"""
desc_paragraph_style = ParagraphStyle(
    'CustomStyle',
    parent=getSampleStyleSheet()['Normal'],
    backColor=colors.white,
    fontName='Helvetica-Bold'
)
desc_paragraph = Paragraph(first_item['Désignation'], desc_paragraph_style)
desc_paragraph.wrapOn(pdf_canvas, card_width, card_height)
desc_paragraph.drawOn(pdf_canvas, 0, (card_height-desc_paragraph.height))
"""
desc_paragraph.drawOn(pdf_canvas, 0, (card_height-desc_paragraph.height))

price_paragraph_style = ParagraphStyle(
    'CustomStyle',
    parent=getSampleStyleSheet()['Normal'],
    alignment=1,
    fontName='Helvetica-Bold',
    fontSize=24
)
price_paragraph = Paragraph(f"{first_item['PV MB']}", price_paragraph_style)
price_paragraph.wrapOn(pdf_canvas, card_width, card_height)
price_paragraph.drawOn(pdf_canvas, 0, price_paragraph.height+(3*mm))

rayon_paragraph_style = ParagraphStyle(
    'CustomStyle',
    parent=getSampleStyleSheet()['Normal'],
    fontName='Helvetica',
    fontSize=16
)
rayon_paragraph = Paragraph(f"{first_item['Code Rayon']}", rayon_paragraph_style)
rayon_paragraph.wrapOn(pdf_canvas, card_width*0.3, card_height)
rayon_paragraph.drawOn(pdf_canvas, 0, price_paragraph.height+(7 *mm)+rayon_paragraph.height)

pdf_canvas.showPage()
pdf_canvas.save()
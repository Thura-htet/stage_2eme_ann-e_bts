import argparse
import csv
import pandas as pd
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

    desc_paragraph.drawOn(pdf_canvas, x, (y + card_height - desc_paragraph.height - price_paragraph.height*2))

    os.remove(barcode_file)

# TODO: factorize code
# TODO: could crop the barcode image
# TODO: read excel files


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
        print()
        generate_product_label(pdf_canvas, product, x, y)

    pdf_canvas.save()


def main(args):
    Stock = read_file('test_stock.xlsm', "Sheet1")
    generate_pdf(Stock, "test_doc.pdf", A4)


if __name__ == "__main__":
    # Create an argument parser
    parser = argparse.ArgumentParser(description="Your program description")

    # Add command-line arguments
    parser.add_argument("--arg1", type=int, help="Description of argument 1")
    parser.add_argument("--arg2", type=str, default="default_value", help="Description of argument 2")

    # Parse the command-line arguments
    args = parser.parse_args()

    # Call the main function with parsed arguments
    main(args)

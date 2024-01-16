import pandas as pd
import os

from PIL import Image

from barcode import Code128
from barcode import EAN13
from barcode.writer import ImageWriter
from barcode.writer import SVGWriter
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import letter, legal, A4, A3

from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph


class PDFProcessor:
    def __init__(self, configs):
        self.stock = None
        self.pdf_canvas = None
        self.out_file = None
        self.num_generated = 0
        self.errors = []

        self.page_size = self.get_page_size(configs["page_size"])
        self.configs = configs

        self.doc_width, self.doc_height = self.page_size
        self.top, self.left = configs['margin']['top'] * mm, configs['margin']['left'] * mm

        hor_space, ver_space = self.doc_width - self.left * 2, self.doc_height - self.top * 2
        self.rows, self.cols = configs['dimensions']['rows'], configs['dimensions']['cols']
        self.card_width, self.card_height = hor_space / self.cols, ver_space / self.rows

    @staticmethod
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

    def read_file(self, file, sheet_name=None, header=0):
        file_ext = file.split(".")[-1]
        expected_columns = ["Référence", "Désignation", "Code Rayon", "PV MB"]

        if file_ext == "csv":
            try:
                with open(file, newline='') as csvfile:
                    df = pd.read_csv(csvfile, sep='\t')
                    self.stock = df.to_dict('records')
                    print(f"read {file} successfully")
            except Exception as e:
                print(f"Error reading {file}: {e}")

        elif file_ext == "xlsm" or file_ext == "xlsx":
            try:
                with pd.ExcelFile(file) as xls:
                    df = pd.read_excel(xls, sheet_name=sheet_name, header=header)
                    print(df.columns)
                    self.stock = df.to_dict('records')
                    print(f"read {file} successfully")
            except Exception as e:
                print(f"Error reading {file}: {e}")

        product_count = len(self.stock)
        page_count = 1 + product_count // (self.configs['dimensions']['rows'] * self.configs['dimensions']['cols'])
        self.out_file = (f"{file}_{self.configs['dimensions']['rows']}x{self.configs['dimensions']['cols']}"
                         f"-de-page-0001-à-page-{page_count:04d}.pdf")
        self.out_file = self.out_file.replace(f'.{file_ext}', '')

    @staticmethod
    def crop_barcode(barcode_file):
        image = Image.open(barcode_file)

        # Convert millimeters to pixels based on image resolution (assuming 300 DPI)
        dpi = 72
        top = int(2 * dpi / 25.4)
        left = int(5 * dpi / 25.4)
        right = int(5 * dpi / 25.4)
        bottom = int(5 * dpi / 25.4)

        cropped_image = image.crop((left, top, image.width - right, image.height - bottom))

        cropped_image.save(barcode_file)

    def generate_product_label(self, pdf_canvas, product, x, y):
        barcode_file = f"barcode_{str(product['Référence']).strip().replace('/', '_')}.png"
        options = {
            'module_height': 12,
            'quiet_zone': 3,
            'font_size': 5,
            'text_distance': 2.3
        }
        writer = ImageWriter()
        writer.set_options(options)
        # barcode = Code128(str(product['Référence']), writer)
        # barcode = Code128(str(product['Référence']), writer)
        # barcode.save(barcode_file, options)

        with open(barcode_file, "wb") as outfile:
            try:
                Code128(str(product['Référence']), writer=ImageWriter()).write(outfile, options)
            except Exception as e:
                print(f"Error generating barcode image for product <{product}>: {e}")

        # self.crop_barcode(barcode_file)

        rayon_paragraph_style = ParagraphStyle(
            'CustomStyle',
            parent=getSampleStyleSheet()['Normal'],
            fontName='Helvetica',
            fontSize=10
        )
        rayon_paragraph = Paragraph(f"{product['Code Rayon']}", rayon_paragraph_style)
        rayon_paragraph.wrapOn(pdf_canvas, self.card_width * 0.3, self.card_height)

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
        price_paragraph.wrapOn(pdf_canvas, self.card_width, self.card_height)

        desc_paragraph_style = ParagraphStyle(
            'CustomStyle',
            parent=getSampleStyleSheet()['Normal'],
            backColor=colors.white,
            fontName='Helvetica-Bold'
        )
        desc_paragraph = Paragraph(product['Désignation'], desc_paragraph_style)
        desc_paragraph.wrapOn(pdf_canvas, self.card_width, self.card_height)

        pdf_canvas.drawInlineImage(
            barcode_file,
            x + (self.card_width * 0.4), y,
            width=self.card_width * 0.6,
            height=self.card_height * 0.8
        )
        rayon_paragraph.drawOn(pdf_canvas, x, y + rayon_paragraph.height)
        price_paragraph.drawOn(pdf_canvas, x, y + self.card_height - price_paragraph.height)
        desc_paragraph.drawOn(
            pdf_canvas,
            x, (y + self.card_height - desc_paragraph.height - price_paragraph.height * 2)
        )

        os.remove(barcode_file)

    def generate_pdf(self, pdf_file):
        sorted_products = sorted(self.stock, key=lambda item: str(item["Code Rayon"]))
        self.pdf_canvas = canvas.Canvas(self.out_file, pagesize=self.page_size)
        y = 0

        for i, product in enumerate(sorted_products):
            if (i != 0) and (i % (self.rows * self.cols) == 0):
                self.pdf_canvas.showPage()

            row = (i // self.cols) % self.rows
            col = i % self.cols
            x = self.left + (col * self.card_width)
            # the +1 is because we are starting from the top of the document
            # and the file is being written from the bottom.
            # we need to save the space for the height of one row
            y = self.doc_height - ((row + 1) * self.card_height) - self.top

            # print(product)
            # print()

            try:
                self.generate_product_label(self.pdf_canvas, product, x, y)
                self.num_generated += 1
            except Exception as e:
                print(f"Error generating product label for product {product['Référence']}: {e}")
                print()
                self.errors.append({
                    "row": i,
                    "Référence": product["Référence"],
                    "Désignation": product["Désignation"],
                    "Code Rayon": product["Code Rayon"],
                })

        self.pdf_canvas.save()

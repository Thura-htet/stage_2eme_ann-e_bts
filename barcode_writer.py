import argparse
from helpers import *
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

from config import *


# TODO: factorize code
# TODO: could crop the barcode image
# TODO: manage configs from here
# TODO: add error handling


def main(args):
    stock = read_file(args.file_path, args.sheet_name)
    generate_pdf(stock, "test_doc.pdf", args.page_size)


if __name__ == "__main__":
    # Create an argument parser
    parser = argparse.ArgumentParser(description="Creates a pdf file containing barcodes from the stock file provided.")

    # Add command-line arguments
    parser.add_argument("file_path", type=str, help="Path to the stock file (.csv or .xlsm/.xlsx format)")
    parser.add_argument("--sheet_name",
                        type=str, default="Sheet1",
                        help="Name of the excel sheet which contains the stock data"
                        )
    parser.add_argument("--page_size",
                        type=str,
                        default="A4",
                        choices=["A4", "A3", "letter", "legal"],
                        help="Page size or type of paper (A4, A3, letter, legal)"
                        )

    # Parse the command-line arguments
    args = parser.parse_args()

    # Call the main function with parsed arguments
    main(args)

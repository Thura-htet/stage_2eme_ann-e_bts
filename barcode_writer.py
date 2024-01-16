import argparse

from helpers import PDFProcessor
from config import *


# TODO: factorize code
# TODO: could crop the barcode image
# TODO: manage configs from here
# TODO: add error handling


def main(args):
    configs = read_config("config.json")  # move this to main and call it before the object initialization

    doc = PDFProcessor(configs)
    doc.read_file(args.file_path, args.sheet_name)
    doc.generate_pdf("test_doc.pdf")


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

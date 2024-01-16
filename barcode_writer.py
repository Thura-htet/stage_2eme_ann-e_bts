import argparse

from helpers import PDFProcessor
from config import *


# TODO: factorize code
# TODO: could crop the barcode image
# TODO: manage configs from here
# TODO: add error handling


def main(args):
    print(args)
    configs = read_config("config.json")  # move this to main and call it before the object initialization

    doc = PDFProcessor(configs)
    doc.read_file(args.file_path, args.sheet_name, (args.header_row-1))
    doc.generate_pdf("test_doc.pdf")

    print(f"number of labels generated: {doc.num_generated}")
    print()
    print(f"number of errors: {len(doc.errors)}")
    print(f"erroneous labels for products in rows: {[product['row'] for product in doc.errors]}"
          f"\nin sheet <{args.sheet_name}> of file <{args.file_path}>")


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
    parser.add_argument("--header_row",
                        type=int,
                        default=0,
                        help="Specify the position of the header row (e.g. 1 if the first row in the sheet is the "
                             "header)")

    # Parse the command-line arguments
    args = parser.parse_args()

    # Call the main function with parsed arguments
    main(args)

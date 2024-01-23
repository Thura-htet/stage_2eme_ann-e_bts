from tkinter import *
from tkinter import ttk, filedialog

import pandas as pd

from config import *
from helpers import PDFProcessor

import helpers


class GUIApp:
    def __init__(self, root):
        self.file_path = None
        self.file_ext = None
        self.spread_sheets = None
        self.selected_sheet = None
        self.xls = None
        self.header_row_index = None
        self.spinbox_variable = IntVar()
        self.data_verified = False
        self.df = None

        self.root = root
        self.root.title("Label Generator")

        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.grid(column=0, row=0, sticky="nsew")

        self.file_label = ttk.Label(self.main_frame, text="No file selected.")
        self.file_label.grid(column=0, row=1, columnspan=2, pady=10)

        self.open_file_button = ttk.Button(self.main_frame, text="Browse", command=self.open_file_dialog, padding=0)
        self.open_file_button.grid(column=3, row=1, columnspan=2, pady=10)

        # reconsider whether these variables need to be class variables
        self.sheet_label = None
        self.sheets_combobox = None
        # self.choose_sheet_button = None

        # self.button_style = ttk.Style()
        # self.button_style.configure('Custom.TButton', state='disabled', background='#999999')
        self.generate_button = ttk.Button(self.main_frame,
                                          text="Generate PDF",
                                          command=self.generate_pdf)
        self.generate_button['state'] = 'disabled'
        self.generate_button.grid(column=5, row=1, columnspan=2, pady=10)

    def open_file_dialog(self):
        self.file_path = filedialog.askopenfile(initialdir="/", title="Select a file").name
        self.file_label.config(text=f"Selected File: {self.file_path}")

        self.file_ext = self.file_path.split(".")[-1]
        if self.file_ext == "xlsm" or self.file_ext == "xlsx":
            self.xls = pd.ExcelFile(self.file_path)
            self.spread_sheets = self.xls.sheet_names

            self.sheet_label = ttk.Label(self.main_frame, text="Choose a spread sheet:")
            self.sheet_label.grid(column=0, row=2, pady=10)
            self.update_sheet_combobox()
            # self.choose_sheet_button = ttk.Button(self.main_frame, text="Choose", command=self.verify_data)
            # self.choose_sheet_button.grid(column=2, row=2)

    # you want to show the corresponding columns every time an item on the combo box has been selected
    def update_sheet_combobox(self):
        if not self.sheets_combobox:
            self.sheets_combobox = ttk.Combobox(self.main_frame, values=self.spread_sheets)
            self.sheets_combobox.grid(column=1, row=2, sticky="nsew")
        self.sheets_combobox["values"] = self.spread_sheets
        self.sheets_combobox.bind("<<ComboboxSelected>>", self.find_header)

    def find_header(self, event):
        self.selected_sheet = self.sheets_combobox.get()
        self.df = pd.read_excel(self.xls, sheet_name=self.selected_sheet, header=self.header_row_index)

        for index, row in self.df.iterrows():
            if self.is_header(row.values):
                self.header_row_index = int(str(index))
                break

        header_row = self.df.iloc[self.header_row_index]
        df = self.df.iloc[self.header_row_index+1:]
        df.columns = header_row

        self.generate_button['state'] = 'normal' if self.header_row_index is not None else 'disabled'

    @staticmethod
    def is_header(row):
        return (all(x_col in row for x_col in ["Référence", "Désignation", "Code Rayon"])
                and any(x_col in row for x_col in ['PV MB', 'PV MB CALCULE COEFF']))

    def generate_pdf(self):
        # load data and close xls here
        print("Generating PDF...")
        configs = read_config("configs.json")

        doc = PDFProcessor(configs)
        doc.read_file(self.file_path, self.selected_sheet, self.header_row_index)
        doc.generate_pdf("test.doc.pdf")


root = Tk()
app = GUIApp(root)
root.mainloop()

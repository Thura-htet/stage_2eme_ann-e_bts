import threading
from tkinter import *
from tkinter import ttk, filedialog
import subprocess

import pandas as pd

from helpers import PDFProcessor

# TODO: progress bar
# TODO: clean up and reorganize the unnecessary object instances
# TODO: cache barcode png
# TODO: clean data from excel
# TODO: open the pdf file after execution
# TODO: start line number
# TODO: add more data


class GUIApp:
    def __init__(self, master):
        self.file_path = StringVar()
        self.file_ext = None
        self.spread_sheets = None
        self.selected_sheet = None
        self.xls = None
        self.header_row_index = None
        self.spinbox_variable = IntVar()
        self.data_verified = False
        self.df = None
        self.out_file = None
        self.configs = {
            'page_size': 'A4',
            'barcode_type': 'Code128',
            'margin': {
                'top': 7,
                'left': 14
            },
            'dimensions': {
                'rows': 7,
                'cols': 3
            }
        }

        self.master = master
        self.master.title("Label Generator")
        self.master.rowconfigure(0, weight=1)
        self.master.columnconfigure(0, weight=1)

        self.main_frame = ttk.Frame(self.master, padding=40)
        self.main_frame.grid(column=0, row=0, sticky="nsew")
        for i in range(3):
            weight = 2 if i == 1 else 1
            self.main_frame.rowconfigure(i, weight=weight)
            self.main_frame.columnconfigure(i, weight=1)

        """
        ===============
        FIRST ROW FRAME
        ===============
        """
        self.first_row_frame = Frame(self.main_frame)
        self.first_row_frame.grid(column=0, row=0, columnspan=3, sticky="nsew")
        for i in range(3):
            self.first_row_frame.columnconfigure(i, weight=1)
        for i in range(4):
            self.first_row_frame.rowconfigure(i, weight=1)

        self.file_label = ttk.Label(self.first_row_frame, text="Select a file:", anchor="e")
        self.file_label.grid(column=0, row=0, sticky="nsew")

        self.file_entry = ttk.Entry(self.first_row_frame, textvariable=self.file_path, state="readonly")
        self.file_entry.grid(column=1, row=0, sticky="ew")
        self.file_error = None

        self.open_file_button = ttk.Button(self.first_row_frame, text="Browse", command=self.open_file_dialog)
        self.open_file_button.grid(column=2, row=0)

        self.sheet_label = None
        self.sheets_combobox = None
        self.header_error = None
        """
        ================
        SECOND ROW FRAME
        ================
        """
        self.second_row_frame = Frame(self.main_frame)
        self.second_row_frame.grid(column=0, row=1, columnspan=3, sticky="nsew")
        # for i in range(2):
        #     self.second_row_frame.columnconfigure(i, weight=1)
        # self.second_row_frame.rowconfigure(0, weight=1)

        self.settings_label = ttk.Label(self.second_row_frame, text="Configurations", anchor="w", width=0)
        self.settings_label.grid(column=0, row=0, sticky="nsew")

        self.page_size_label = ttk.Label(self.second_row_frame, text="page size:", anchor="e", width=0)
        self.page_size_label.grid(column=1, row=1, sticky="e")
        self.page_size = StringVar(value=self.configs['page_size'])
        self.page_size_combobox = ttk.Combobox(self.second_row_frame, textvariable=self.page_size,
                                               values=["A4", "A3", "Letter", "Legal"])
        self.page_size_combobox.grid(column=2, row=1, sticky="nsew")

        self.barcode_type_label = ttk.Label(self.second_row_frame, text="barcode type", anchor="w", width=0)
        self.barcode_type_label.grid(column=4, row=1, sticky="nsew")
        self.barcode_type = StringVar(value=self.configs['barcode_type'])
        self.barcode_type_combobox = ttk.Combobox(self.second_row_frame, textvariable=self.barcode_type,
                                                  values=["EAN13", "Code128"])
        self.barcode_type_combobox.grid(column=5, row=1, sticky="nsew")

        self.margins_label = ttk.Label(self.second_row_frame, text="margins", anchor="e", width=0)
        self.margins_label.grid(column=0, row=2, sticky="nsew")
        self.ver_label = ttk.Label(self.second_row_frame, text="top:", anchor="e", width=0)
        self.ver_label.grid(column=1, row=3, sticky="e")
        self.ver_space = IntVar(value=self.configs['margin']['top'])
        self.ver_entry = ttk.Entry(self.second_row_frame, textvariable=self.ver_space, width=5)
        self.ver_entry.grid(column=2, row=3, sticky="w")

        self.hor_label = ttk.Label(self.second_row_frame, text="left:", anchor="e", width=0)
        self.hor_label.grid(column=1, row=4, sticky="e")
        self.hor_space = IntVar(value=self.configs['margin']['left'])
        self.hor_entry = ttk.Entry(self.second_row_frame, textvariable=self.hor_space, width=5)
        self.hor_entry.grid(column=2, row=4, sticky="w")

        self.dimensions_label = ttk.Label(self.second_row_frame, text="dimensions", anchor="e", width=0)
        self.dimensions_label.grid(column=0, row=5, sticky="nsew")

        self.cols_label = ttk.Label(self.second_row_frame, text="columns:", anchor="e", width=0)
        self.cols_label.grid(column=1, row=6, sticky="e")
        self.cols = IntVar(value=self.configs['dimensions']['cols'])
        self.cols_entry = ttk.Entry(self.second_row_frame, textvariable=self.cols, width=5)
        self.cols_entry.grid(column=2, row=6, sticky="w")

        self.rows_label = ttk.Label(self.second_row_frame, text="rows:", anchor="e", width=0)
        self.rows_label.grid(column=1, row=7, sticky="e")
        self.rows = IntVar(value=self.configs['dimensions']['rows'])
        self.rows_entry = ttk.Entry(self.second_row_frame, textvariable=self.rows, width=5)
        self.rows_entry.grid(column=2, row=7, sticky="w")

        self.apply_button = ttk.Button(self.second_row_frame,
                                       text="Apply",
                                       command=self.apply_configs)
        self.apply_button.grid(column=5, row=8)
        """
        ===============
        THIRD ROW FRAME
        ===============
        """
        self.third_row_frame = Frame(self.main_frame)
        self.third_row_frame.grid(column=0, row=2, columnspan=3, sticky="nsew")
        for i in range(3):
            self.third_row_frame.columnconfigure(i, weight=1)
        self.third_row_frame.rowconfigure(0, weight=1)

        # self.creating_label = ttk.Label(self.third_row_frame, font=("Arial", 10, "normal"), foreground="grey")
        # self.creating_label.grid(column=0, row=0, sticky="w")

        # self.progressbar = ttk.Progressbar(self.third_row_frame, orient="horizontal", mode="determinate")
        # self.progressbar.grid(column=0, row=1, sticky="nsew")

        self.generate_button = ttk.Button(self.third_row_frame,
                                          text="Generate PDF",
                                          command=self.generate_pdf)
        self.generate_button['state'] = 'disabled'
        self.generate_button.grid(column=2, row=0, sticky='e')

    def open_file_dialog(self):
        self.file_path.set(filedialog.askopenfile(initialdir="/", title="Select a file").name)
        # self.file_label.config(text=f"Selected File: {self.file_path}")

        self.file_ext = self.file_path.get().split(".")[-1]
        if self.file_ext == "xlsm" or self.file_ext == "xlsx":
            if self.file_error:
                self.file_error.destroy()
            self.xls = pd.ExcelFile(self.file_path.get())
            self.spread_sheets = self.xls.sheet_names
            # TODO: the sheet label is probably distorting the column widths
            self.sheet_label = ttk.Label(self.first_row_frame, text="Choose spreadsheet:", anchor="e")
            self.sheet_label.grid(column=0, row=2, sticky="e")
            self.update_sheet_combobox()
        elif self.file_ext == "csv":
            if self.file_error:
                self.file_error.destroy()
            self.generate_button['state'] = 'normal'
            self.header_row_index = 0
            # self.choose_sheet_button = ttk.Button(self.main_frame, text="Choose", command=self.verify_data)
            # self.choose_sheet_button.grid(column=2, row=2)
        else:
            self.file_error = ttk.Label(self.first_row_frame, font=("Arial", 10, "bold"), foreground="red",
                                        wraplength="500",
                                        text=f"The type of the file you have chosen "
                                             f"<{self.file_path.get().split('/')[-1]}> is not supported. "
                                             f"Please choose a .csv or a .xslm/.xlsx file.")
            self.file_error.grid(column=0, row=1, columnspan=3, sticky="ew")
            if self.sheets_combobox:
                self.sheets_combobox.destroy()
                self.sheet_label.destroy()

    # you want to show the corresponding columns every time an item on the combo box has been selected
    def update_sheet_combobox(self):
        if not self.sheets_combobox:
            self.sheets_combobox = ttk.Combobox(self.first_row_frame, values=self.spread_sheets)
            self.sheets_combobox.grid(column=1, row=2, sticky="w")
        self.sheets_combobox["values"] = self.spread_sheets
        self.sheets_combobox.bind("<<ComboboxSelected>>", self.find_header)

    def find_header(self, event):
        if self.header_error:
            self.header_error.destroy()
        self.header_row_index = None
        self.selected_sheet = self.sheets_combobox.get()
        self.df = pd.read_excel(self.xls, sheet_name=self.selected_sheet, header=self.header_row_index)

        for index, row in self.df.iterrows():
            if self.is_header(row.values):
                self.header_row_index = int(str(index))
                break

        print(self.header_row_index)
        if self.header_row_index is not None:
            header_row = self.df.iloc[self.header_row_index]
            df = self.df.iloc[self.header_row_index + 1:]
            df.columns = header_row
            self.df = df
            self.generate_button['state'] = 'normal'
        else:
            self.generate_button['state'] = 'disabled'
            self.header_error = ttk.Label(self.first_row_frame, font=("Arial", 10, "bold"), foreground="red",
                                          wraplength="500",
                                          text=f"The spreadsheet \"{self.selected_sheet}\" doesn't contain "
                                               f"necessary information to create a shelf label. "
                                               f"Please try another spreadsheet.")
            self.header_error.grid(column=0, row=3, columnspan=3, sticky="ew")

    @staticmethod
    def is_header(row):
        return (all(x_col in row for x_col in ["Référence", "Désignation", "Code Rayon"])
                and any(x_col in row for x_col in ['PV MB', 'PV MB CALCULE COEFF']))

    def apply_configs(self):
        self.configs["page_size"] = self.page_size.get()
        self.configs["barcode_type"] = self.barcode_type.get()
        self.configs['dimensions']['rows'] = self.rows.get()
        self.configs['dimensions']['cols'] = self.cols.get()
        self.configs['margin']['top'] = self.ver_space.get()
        self.configs['margin']['left'] = self.hor_space.get()
        print(self.configs)

    def generate_pdf(self):
        # load data and close xls here
        # creating_label = ttk.Label(self.third_row_frame, text=f"Creating label: ", font=("Arial", 10, "normal"),
        #                            foreground="grey")
        # creating_label.grid(column=0, row=0, sticky="w")

        in_file = self.file_path.get().split('/')[-1]
        self.out_file = filedialog.askdirectory(initialdir=self.file_path.get().replace("in_file", ""),
                                                title='Select folder')
        if self.out_file:
            print("Generating PDF...")
            self.generate_button['text'] = 'Generating...'
            self.generate_button['state'] = 'disabled'

            doc = PDFProcessor(self.configs, self.df)
            self.xls.close()

            num_products = doc.num_products  # TODO: implement getter, setter
            page_count = 1 + num_products // (self.configs['dimensions']['rows'] * self.configs['dimensions']['cols'])
            self.out_file = (f"{self.out_file}/{self.file_path.get().split('/')[-1]}"
                             f"_{self.configs['dimensions']['rows']}x{self.configs['dimensions']['cols']}"
                             f"-de-page-0001-à-page-{page_count:04d}.pdf")
            self.out_file = self.out_file.replace(f'.{self.file_ext}', '')
            # doc.read_file(self.file_path, self.selected_sheet, self.header_row_index)

            # could try making the thread into a function
            # self.creating_label.config(text=f"Creating label: ")

            thread = threading.Thread(target=doc.generate_pdf, args=(self.out_file,))
            # self.progressbar.configure(maximum=doc.num_products)
            # self.progressbar.start()
            # self.third_row_frame.update_idletasks()
            thread.start()

            current_product_row = 0
            while thread.is_alive():
                # need to pass a condition to check whether there's been an update
                doc.event.wait()
                num_generated, product_generated = doc.get_progress_info()
                if product_generated is not None:
                    print(f"thread! creating shelf label for product: {product_generated['Référence']}")
                    # self.creating_label.configure(text=f"Creating label for: {product_generated['Référence']}")
                    # creating_label.configure(text=f"Creating label for: {product_generated['Référence']}")
                    # self.progressbar.step(1)
                    # self.progressbar.update_idletasks()
                    doc.event.clear()

            thread.join()
            print("program terminated")
            # TODO: apparently need the additional option shell = True for windows
            subprocess.run(['open', self.out_file], check=True)
            self.generate_button['state'] = 'normal'
            self.generate_button['text'] = 'Generate PDF'


root = Tk()
root.geometry('750x550')
root.resizable(width=False, height=False)

app = GUIApp(root)
root.mainloop()

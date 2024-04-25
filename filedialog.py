import tkinter as tk
import tkinter.filedialog
from tkinter import messagebox

def select_excel_file():

    # open file 
    file_path = tkinter.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])

    if file_path:
            return file_path
    else:
        messagebox.showinfo("Notice", "Please select a excel file.")
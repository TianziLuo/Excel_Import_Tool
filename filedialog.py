import tkinter.filedialog
from tkinter import filedialog, messagebox

def select_excel_file():

    # open file 
    file_open_path = tkinter.filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsx")])

    if file_open_path:
            return file_open_path
    else:
        messagebox.showinfo("Notice", "Please select a excel file.")


def save_excel_file():
    file_save_path = filedialog.asksaveasfilename(initialdir="/", title="Save File As", filetypes=[("Excel files", "*.xlsx")])
    
    if not file_save_path:
        messagebox.showinfo("Notice", "Please select a location to save the XLSX file.")
        return None

    if "." not in file_save_path:
        file_save_path += ".xlsx"
        return file_save_path

## print(save_excel_file())
import tkinter as tk
from tkinter import filedialog
from pdf2docx import Converter

def convert_pdf_to_word():
    root = tk.Tk()
    root.withdraw()

    pdf_file = filedialog.askopenfilename(title = "Select PDF file", filetypes = (("PDF Files","*.pdf"),))

    word_file = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=(("Word Document", "*.docx"),("All Files", "*.*")))

    cv = Converter(pdf_file)

    cv.convert(word_file, start=0, end=None)

    cv.close()

convert_pdf_to_word()
import tkinter as tk
from tkinter import filedialog
from docx2pdf import convert

def convert_word_to_pdf():
    root = tk.Tk()
    root.withdraw()

    word_file = filedialog.askopenfilename(title = "Select Word file", filetypes = (("Word Files","*.docx"),))

    pdf_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=(("PDF Files", "*.pdf"),("All Files", "*.*")))

    convert(word_file, pdf_file)

convert_word_to_pdf()
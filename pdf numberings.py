import os
import tkinter as tk
from tkinter import filedialog

from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

def create_page_pdf(num, tmp):
    c = canvas.Canvas(tmp)
    for i in range(1, num + 1):
        c.drawString((210 // 2) * mm, (20) * mm, str(i))
        c.showPage()
    c.save()

def add_page_numgers():

    tmp = "__tmp.pdf"

    root = tk.Tk()
    root.withdraw()

    pdf_path = filedialog.askopenfilename(title = "Select PDF file", filetypes = (("PDF Files","*.pdf"),))

    newpath = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=(("PDF Files", "*.pdf"),("All Files", "*.*")))

    writer = PdfWriter()
    with open(pdf_path, "rb") as f:
        reader = PdfReader(f)
        n = len(reader.pages)

        create_page_pdf(n, tmp)

        with open(tmp, "rb") as ftmp:
            number_pdf = PdfReader(ftmp)
            for p in range(n):
                page = reader.pages[p]
                number_layer = number_pdf.pages[p]
                page.merge_page(number_layer)
                writer.add_page(page)

            if len(writer.pages) > 0:
                with open(newpath, "wb") as f:
                    writer.write(f)
        os.remove(tmp)

add_page_numgers()
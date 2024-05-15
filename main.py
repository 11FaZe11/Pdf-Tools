import os
import tkinter as tk
from tkinter import filedialog
from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from docx2pdf import convert
from pdf2docx import Converter
import time
## Run in your terminal >>>> pip install PyPDF2 docx2pdf pdf2docx 


while True:
  try:
    print("""  
1- Convert PDF to WORD file.
2- Convert WORD to PDF file.
3- Add numbers on PDF file.
0- End the program.
""")

    first_choice = int(input("\nEnter the operation you wish..... \n"))


    if first_choice in range(0,4):

        if first_choice == 1:
            print("""Welcome to the PDF to WORD converter""")
            def convert_pdf_to_word():
                root = tk.Tk()
                root.withdraw()

                pdf_file = filedialog.askopenfilename(title="Select PDF file", filetypes=(("PDF Files", "*.pdf"),))

                word_file = filedialog.asksaveasfilename(defaultextension=".docx",
                                                         filetypes=(("Word Document", "*.docx"), ("All Files", "*.*")))

                cv = Converter(pdf_file)

                cv.convert(word_file, start=0, end=None)

                cv.close()


            convert_pdf_to_word()





        elif first_choice == 2:
            print("""Welcome to the WORD to PDF converter""")
            def convert_word_to_pdf():
                root = tk.Tk()
                root.withdraw()

                word_file = filedialog.askopenfilename(title="Select Word file", filetypes=(("Word Files", "*.docx"),))

                pdf_file = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                        filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*")))

                convert(word_file, pdf_file)


            convert_word_to_pdf()





        elif first_choice == 3:
            print("""Welcome to the PDF numbering tool""")
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

                pdf_path = filedialog.askopenfilename(title="Select PDF file", filetypes=(("PDF Files", "*.pdf"),))

                newpath = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                       filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*")))

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


        elif first_choice == 0:
            break

    else:
      print("Invalid input. Please enter from the previos choices")
    print("""
        the program will start in 3 seconds
        """)
    for abc in range(1,2):
        print(abc)
        time.sleep(1)

  except ValueError:

    print("Invalid input. Please enter a number.")



  if first_choice in range(0,1000000000000000000000000000000000000000000000000) and first_choice != 0:
    print("Do you want to enter another option? (yes/no)")

    answer = input().lower()
    if answer != 'yes':
      break
print("Thank you for using this tool !!")

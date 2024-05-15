import os
import time
import tkinter as tk
from tkinter import *
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from docx2pdf import convert
from pdf2docx import Converter
from tkinter import filedialog, messagebox
from PyPDF2 import PdfFileReader, PdfReader, PdfWriter



while True:
  try:
    print("""  
1- Convert PDF to WORD file.
2- Convert WORD to PDF file.
3- Add numbers on PDF file.
4- PDF locker tool.
0- End the program.
""")

    first_choice = int(input("\nEnter the operation you wish..... \n"))


    if first_choice in range(0,10):

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



        elif first_choice == 4:
            print("""Welcome to the PDF locker tool""")
            def select_source_file():
                file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
                if file_path:
                    source.set(file_path)
                    target.set(file_path[:-4] + "_locked.pdf")


            def protect_pdf():
                input_pdf_path = source.get()
                output_pdf_path = target.get()

                if not input_pdf_path or not output_pdf_path:
                    messagebox.showerror("Error", "Please select source and target PDF files.")
                    return

                password = password_entry.get()
                if not password:
                    messagebox.showerror("Error", "Please enter a password.")
                    return

                pdf_writer = PdfWriter()
                pdf_reader = PdfReader(input_pdf_path)

                for page_num in range(len(pdf_reader.pages)):
                    pdf_writer.add_page(pdf_reader.pages[page_num])

                pdf_writer.encrypt(user_pwd=password, use_128bit=True)

                with open(output_pdf_path, 'wb') as encrypted_pdf_file:
                    pdf_writer.write(encrypted_pdf_file)

                messagebox.showinfo("Success", "PDF encrypted and saved successfully.")


            root = Tk()
            root.title("PDF Locker App")
            root.geometry("700x400")
            root.resizable(True, True)

            frame = Frame(root)
            frame.pack(padx=20, pady=20)

            source = StringVar()
            target = StringVar()

            Label(frame, text="Source File PDF:", font="arial 10 bold").grid(row=0, column=0, sticky=W)
            entry1 = Entry(frame, width=30, textvariable=source, font='arial 15', bd=1)
            entry1.grid(row=0, column=1)
            Button(frame, text="Select", command=select_source_file).grid(row=0, column=2)

            Label(frame, text="Target PDF file:", font="arial 10 bold").grid(row=1, column=0, sticky=W)
            entry2 = Entry(frame, width=30, textvariable=target, font='arial 15', bd=1)
            entry2.grid(row=1, column=1)

            Label(frame, text="Enter Password:", font="arial 10 bold").grid(row=2, column=0, sticky=W)
            password_entry = Entry(frame, width=30, font='arial 15', bd=1)
            password_entry.grid(row=2, column=1)

            Button(frame, text="Lock PDF", command=protect_pdf).grid(row=3, column=1, pady=20)

            root.mainloop()



        elif first_choice == 0:
            break

    else:
      print("Invalid input. Please enter from the previos choices")
    print("""
        the program will start in 3 seconds
        """)
    for abc in range(1,3):
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

# Import libraries
import os
import time
import pandas as pd
import pdfkit

# Insert the directory path in here
from PyPDF2 import PdfFileMerger
from docx2pdf import convert

path = 'C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets'


def wordtopdf():
    convert('C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\8. Index V, Comp, Form 26AS.docx',
            'C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\converted.pdf')


def exceltopdf():
    df = pd.read_excel(
        "C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\Computation of Income AY 2021-22.xlsx")
    df.to_html("C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\file.html")
    pdfkit.from_file("C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\file.html",
                     "C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\file.pdf")


def mergeallpdf():
    pdfs = ['C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\converted.pdf',
            'C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\file.pdf',
            'C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\Form 26AS AY 2021-22 Rama Mohan Kavoori.pdf',
            'C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\ITR V AY 2021-22 Rama Mohan.pdf']

    merger = PdfFileMerger()

    for pdf in pdfs:
        merger.append(pdf)

    merger.write("C:\\Users\\atchu\\PycharmProjects\\pythonProject\\assets\\result.pdf")
    merger.close()


wordtopdf()

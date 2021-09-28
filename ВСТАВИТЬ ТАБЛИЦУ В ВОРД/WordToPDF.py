"""We will convert .docx docs to PDF"""
import os
from tkinter import Tk
from tkinter.filedialog import askdirectory
# turn off tkinter because we dont need the whole module to work
Tk().withdraw()
# choose folder
file_directory = askdirectory()
import win32com.client

def change_file_name(file_directory):
    """Lets change symbols in path to make win32com.Doc works"""
    right_name = ''
    for i in range(len(file_directory)):
        if file_directory[i] == "/":
            right_name += '\\'
        else:
            right_name += file_directory[i]
    return right_name

file_directory = change_file_name(file_directory)

def convertation_to_pdf(file_directory):
    """Convert docx to pdf"""
    word = win32com.client.Dispatch('Word.Application')
    # needed format to convert
    wdFormatPDF = 17
    for filename in os.listdir(file_directory):
        if filename.endswith(".docx"):
            doc = word.Documents.Open(f'{file_directory}\\{filename}')
            filename = filename.rstrip(".docx")
            doc.SaveAs(f'{file_directory}\\{filename}_out_file',
                       FileFormat=wdFormatPDF)
            doc.Close()
        else:
            continue

    return print('all right!')

convertation_to_pdf(file_directory)

"""Split pdf doc by pages"""
from PyPDF2 import PdfFileWriter, PdfFileReader
# define input file
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory
# Используем библиотеку Tkinter чтобы сделать диалоговое окно
# для поиска имени и полного адреса файла
# Нам не нужен полноценный GUI, поэтому отключаем
Tk().withdraw()
# находим путь и сам нужный файл
filename = askopenfilename()
# выбираем директорию, куда будем сохранять файлы
file_directory = askdirectory()
input_pdf = PdfFileReader(filename)
output = PdfFileWriter()
count_pages = input_pdf.getNumPages()
for page in range(count_pages):
    output = PdfFileWriter()
    output.addPage(input_pdf.getPage(page))
    with open(f"{file_directory}/{page}_page.pdf", "wb") as output_stream:
        output.write(output_stream)






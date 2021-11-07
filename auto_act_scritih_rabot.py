"""Создаем АОСР. Работаем с пандой для получения словаря значений из иксель таблицы."""
from pandas import *
# библиотека нужна, чтобы заполнить шаблон из Ворда
from docxtpl import DocxTemplate
import datetime
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
# Теперь выбираем путь и название файла заготовки
filename2 = askopenfilename()
# парсим пандой
# Наш документ открываем как ExcelFile
xls = ExcelFile(filename)
# вводим название файла для сохранения
name = input('Введите название файла:   ')
# парсим пандой на листе с именем "static"
data = xls.parse('static')
# а теперь создаем словарь из значений первого столбца как ключей и остальных столбцов - значений
context_first_part = {}
# счетчик для прохождения по рядам
for n, f in zip(data['keys'], range(data.shape[0])):
    # проверим, если значение равно "nan", то изменим его на " "
    a = str(data['static'][f])
    if str(a) == 'nan':
        context_first_part[n] = ' '
    else:
        context_first_part[n] = data['static'][f]
# парсим пандой на листе с именем "dynamic"
data = xls.parse('dynamic')
print(range(data.shape[0]))
# а теперь создаем словари из значений динамической таблицы
for i in range(1, len(data.columns)):
    context_second_part = {}
    for m, k in zip(range(data.shape[0]), data['Наименование']):
        # проверим, если значение равно "nan", то изменим его на " "
        a = data[data.columns[i]][m]
        # убираем Nan, который добавляется в базу данных, если ячейка
        # пустая
        if str(a) == 'nan':
            context_second_part[k] = " "
        # если это дата, то изменим форматирование на стандартное
        elif isinstance(a, datetime.datetime):
            a = a.strftime("%d.%m.%Y")
            context_second_part[k] = a
        else:
            context_second_part[k] = data[data.columns[i]][m]

    # обновляем словари - соединяем вместе
    context_first_part.update(context_second_part)
    # загружаем заготовку
    template = DocxTemplate(filename2)
    # обрабатываем заготовку
    template.render(context_first_part)
    # сохраняем
    template.save(
        f'{file_directory}/{name}_{i}.docx'
    )

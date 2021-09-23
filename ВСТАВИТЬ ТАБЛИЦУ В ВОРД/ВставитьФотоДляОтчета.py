# For reading file name
import glob
# попробуем работать с подушкой
from PIL import Image
from datetime import datetime
from docx import Document
from docx.shared import Cm
import os

print(glob.glob('.'))  # распечатаем все названия файлов в директории

myListWithoutSlash = []  # создадим список из фотографий
for file_name in glob.iglob('./imgs/*.*', recursive=True):
    print(file_name)
    file_name = file_name.split('\\')[1]  # убираем слэши в названии и берем только последнее
    myListWithoutSlash.append(file_name)
document = Document()  # создаем объект
table = document.add_table(rows=1, cols=3)  # добавляе таблицу
table.style = 'Table Grid'
# дадим название заголовкам
hdr_cells = table.rows[0].cells
hdr_cells[0].text = 'Номер'
hdr_cells[1].text = 'Описание'
hdr_cells[2].text = 'Фотография'
# создадим размер, до которого pillow будет уменьшать фото
size = (700, 700)
n = 1  # счетчик
for i in range(len(myListWithoutSlash)):
    print(i)
    row_cells = table.add_row().cells  # добавляем строку к табилце
    row_cells[0].text = f"фото {n}"  # содержание первого столбца
    picture = Image.open(f'./imgs/{myListWithoutSlash[i]}')  # открываем картинку
    picture.thumbnail(size)  # изменяем размер
    text = ''
    flag = False
    t = 0
    for s in range(len(myListWithoutSlash[i])):
        g = myListWithoutSlash[i][s - 1 - 2 * t]
        if myListWithoutSlash[i][s - 1 -  2 * t] == '.':
            flag = True
        if flag:
            text += myListWithoutSlash[i][s - 1 - 2 * t]
        t += 1
    text = text[::-1]
    picture.save(f'./imgs/{text}_{n}.jpeg')  # пересохраняем в jpeg
    row_cells[1].text = text  # содержание второго столбца
    p = row_cells[2].add_paragraph()
    r = p.add_run()
    r.add_picture(f'./imgs/{text}_{n}.jpeg', width=Cm(5.0), height=Cm(5))
    picture.close()
    os.remove(f'./imgs/{myListWithoutSlash[i]}')  # удаляем старое фото
    n += 1
# для названия файле используем datetime
id_cur = datetime.now().strftime('%d_%m_%y__%H%M')
document.save(f'./output/готов_{id_cur}.docx')

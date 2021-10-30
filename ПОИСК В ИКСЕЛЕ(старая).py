"""Работаем с пандой для поиска нужных данных в таблице иксель."""
import pandas as pd
from docxtpl import DocxTemplate
import datetime
from tkinter import Tk
from tkinter.filedialog import askopenfilename, askdirectory
# Используем библиотеку Tkinter чтобы сделать диалоговое окно
# для выбора файла
# Нам не нужен полноценный GUI, поэтому отключаем
Tk().withdraw()
# находим путь и сам нужный файл
filename = askopenfilename()
# выбираем директорию, куда будем сохранять файлы
file_directory = askdirectory()
# введем строку - что ищем, и преобразуем в словарь,
# искать будем несколько вариантов
print('Введите слова/часть слова - что хотите найти (через пробел)')
w_find = (input().lower())
w_find = w_find.split()
# парсим пандой все листы, при этом опционально пишем sheet_name = None, тогда
# панда будет брать информацию и создавать базу данных из всех листов
# файл должен быть в одной папке со скриптом
sheets = pd.read_excel(filename, sheet_name=None)
# делаем список из названий листов файла
sheets_names = list(sheets)
# создаем список значений из базы данных панды
db_excel = []
# будем вести поиск в нужном столбце базы данных
print('В каком столбце искать? (нумерация с первого) ')
nomer_stolbca_poiska = int(input())
# идем через каждый лист и ищем нужную нам информацию
for name in sheets_names:
    df = ''  # сбрасываем данные из панды
    # вот тут мы опять считываем данные из каждого листа поотдельности
    df = pd.read_excel(filename, sheet_name=name)
    # счетчик, используется для определения номера строки, который добавляем в словарь
    n = 0
    # таким образом методом .iloc[:, столбец] мы забираем данные
    razdel = df.iloc[7, 0]
    for i in df.iloc[:, nomer_stolbca_poiska - 1]:
        # если строчка из этой ячейки содержит нужный нам набор символов,
        # то мы добавляем строку в словарь db_excel
        # теперь работаем со словарем введенных данных
        for word in w_find:
            if str(i).lower().find(word) >= 0:
                # тут видно, что добавляем только начиная с 1 по 10-столбец n-ой строки
                db_excel.append(df.loc[n][1:11])
                db_excel[-1]['Наименование раздела'] = razdel
                for i in range(1,10):
                    # этот участок кода для поиска стоимости работ
                    if df.loc[n + i][2] == 'Всего по позиции':
                        a = df.loc[n+i][9:10]
                        db_excel[-1]['СТОИМОСТЬ'] = a['Unnamed: 9']
                        db_excel[-1]['РАЗДЕЛ'] = name
                # если нашел совпадение, то для исключения задвоения выходим из цикла
                break
        n += 1
# для сохранения в иксель, используя панду переводим словарь в базу данных Панды
df = pd.DataFrame(db_excel)
# создадим уникальное название
u_name = datetime.datetime.now().strftime('%d_%m_%y__%H%M%S')
# сохраняем в иксель
df.to_excel(fr'{file_directory}/{w_find}_{u_name}.xlsx', index = False)
# сообщим, что все готово
print('готово')


"""Работаем с пандой для поиска нужных данных в таблице иксель."""
import pandas as pd
from docxtpl import DocxTemplate
import datetime

df = pd.read_excel(r'D:/ЯНДЕКС ДИСК/YandexDisk/1. МС_ПРО_И_CANDP/_ФРЯЗИНО/Заявки на материалы/ЗАКАЗ_АПС_А3В5_АРГС_МФМ.xlsx')
u_name = datetime.datetime.now().strftime('%d_%m_%y__%H%M%S')
df['All'] = df.groupby(["Код продукции"])['Количество'].transform('sum')
df1 = df.drop_duplicates('Код продукции', keep='first')
df1.to_excel(
    fr'D:/ЯНДЕКС ДИСК/YandexDisk/1. МС_ПРО_И_CANDP/_ФРЯЗИНО/Заявки на материалы/ЗАКАЗ_АПС_А3В5_АРГС_МФМ{u_name}.xlsx',
    index = False)

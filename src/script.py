# Импортируем библиотеки
import openpyxl as xl
import pandas as pd
import re
from pathlib import Path
from os import path
import datetime

def processFile(excel_file_path, format_file, sheet="Sheet1"):

    # Определяем путь к файлу 
    excel_file_path = Path(excel_file_path)
    dc_format_path = Path(format_file)

    # Открываем файл Excel
    file = xl.load_workbook(excel_file_path)

    # Выбираем рабочий лист
    sheet = file[sheet]

    # Определяем даты на которые имеется сток
    dates_dict = {}
    for cell in sheet[3]:
        if re.match(r'\d+\.\d+\.\d+', str(cell.value)):
            dates_dict[cell.value] = {}

    dates_array = []
    for date in dates_dict:
        dates_array.append(date)

    # Определяем неделю на которую у нас файл
    week = 0
    for cell in sheet[2]:
        if re.match(r'\d+', str(cell.value)):
            week = cell.value

    # Определяем год 
    year = 0
    for cell in sheet[1]:
        if re.match(r'\d+', str(cell.value)):
            year = cell.value

    # Создаем датафрейм
    df1 = pd.DataFrame(sheet.values)

    # Избавляемся от ненужных строк
    df1 = df1.drop([0,1,2])

    # Используем теперь первую строку в качестве заголовков
    new_header = df1.iloc[0] 
    df1 = df1[1:] 
    df1.columns = new_header 

    # Переименовываем столбцы с остатками в даты, на которые эти остатки были зафиксированы
    new_header = []
    i = 0
    for name in df1.columns:
        if name == "Остаток":
            new_header.append(f"{dates_array[i]}")
            i += 1
            continue
        new_header.append(name)
    df1.columns = new_header

    # Удаляем ненужные столбцы
    df1 = df1.drop(["Страховой запас", "Блокировка"], axis=1)

    # Объединяем данные по стокам в одном столбце
    # Берем строку
    for index, row in df1.iterrows():
        # Создаем переменную stock - в ней будет храниться финальное значения стока для этой строки
        stock = row[dates_array[0]]
        # Переменная для отобращения к следующей колонке, если в теущей не будет данных
        dArrIdx = 1
        # Если в первой колонке не было данных, то идем искать в следующих колонках
        if stock == "-":
            # Пока в переменной не будет цифры, будем идти от колонки к колонке
            while stock == "-":
                # Как только находим цифру, перезаписываем значение в первой колонке и меняем переменную stock, 
                # чтобы выйти из цикла
                if row[dates_array[dArrIdx]] != "-":
                    row[dates_array[0]] = row[dates_array[dArrIdx]]
                    stock = row[dates_array[dArrIdx]]
                # Если найти в столбце не удалось, то переходим к следующему
                dArrIdx += 1

    # Определяем ненужные столбцы, данные из которых в нужных местах уже перенесены в первый столбец со стоком
    columns_to_delete = dates_array[1:]

    # Удаляем столбцы определенные на предыдущем шаге
    df1 = df1.drop(columns_to_delete, axis=1)

    # Вставляем столбец Год и Неделя, которые берем из исходного файла в начале скрипта
    df1.insert(0, "Неделя", [week]*(len(df1.index)))
    df1.insert(0, "Год", [year]*(len(df1.index)))

    # Переименовываем столбец с остатками из даты в "Остаток"
    df1 = df1.rename(columns={str(dates_array[0]): "Остаток, шт"})

    # Подключаемся к файлу с форматами РЦ
    format_df = pd.read_excel(dc_format_path)
    # На всякий случай меняем формат кода склада на числовой
    df1 = df1.astype({'Код склада': int})
    # Подтягиваем к основному фрейму форматы РЦ на основе кода РЦ Х5
    df1 = df1.merge(format_df, left_on='Код склада', right_on='DC X5 code', how='left')
    # Если не удалось найти сопоставление по кодам, то пишем, что код надо добавить в файл
    df1.fillna(value={'Формат': ""}, inplace=True)
    # Удаляем ненужные столбцы
    df1.drop(columns=['Ship-to', 'DC X5 code', 'DC'], inplace=True)
    # Меняем порядок столбцов
    df1 = df1[['Год', 'Неделя',	'Код склада', 'Наименование точки',	'Формат', 'PLU', 'ШК', 'Наименование Товара', 'Остаток, шт']]

    # Сохраняем файл
    fileName = f"X5_Stock_{week + str(datetime.datetime.now().year)[2:]}.xlsx"
    whereToPut = path.join(excel_file_path.parent, fileName)
    df1.to_excel(whereToPut, index=False)

    return fileName
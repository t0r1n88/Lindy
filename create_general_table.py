import pandas as pd
import openpyxl
import os
from datetime import date


def calculate_age(born):
    """
    Функция для расчета текущего возраста взято с https://stackoverflow.com/questions/2217488/age-from-birthdate-in-python/9754466#9754466
    :param born: дата рождения
    :return: возраст
    """
    today = date.today()
    return today.year - born.year - ((today.month, today.day) < (born.month, born.day))



def update_spreadsheet(path: str, _df, starcol: int = 1, startrow: int = 1, sheet_name: str = "ToUpdate"):
    '''

    :param path: Путь до файла Excel
    :param _df: Датафрейм Pandas для записи
    :param starcol: Стартовая колонка в таблице листа Excel, куда буду писать данные
    :param startrow: Стартовая строка в таблице листа Excel, куда буду писать данные
    :param sheet_name: Имя листа в таблице Excel, куда буду писать данные
    :return:
    '''
    wb = openpyxl.load_workbook(path)
    for ir in range(0, len(_df)):
        for ic in range(0, len(_df.iloc[ir])):
            wb[sheet_name].cell(startrow + ir, starcol + ic).value = _df.iloc[ir][ic]
    wb.save('data/Общая таблица.xlsx')

path = 'data/Общая таблица/'

base_file = 'data/Форма Базы данных от 01.02.2022.xlsx'

# Создаем 2 датафрейма,считывая колонки из файлов
df_dpo = pd.read_excel('data/колонки ДПО.xlsx')
df_po = pd.read_excel('data/колонки ПО.xlsx')

for file in os.listdir(path):
    #Считываем файлы с данными
    # Создаем промежуточный датафрейм с данными с листа ДПО
    temp_dpo = pd.read_excel(f'{path}/{file}',sheet_name='ДПО')
    # Создаем промежуточный датафрейм с данными с листа ПО
    temp_po = pd.read_excel(f'{path}/{file}',sheet_name='ПО')
    # Добавляем промежуточные датафреймы в исходные
    df_dpo = df_dpo.append(temp_dpo,ignore_index=True)
    df_po = df_po.append(temp_po,ignore_index=True)

# Добавляем в датафреймы колонки с текущим возрастом и категорией
df_dpo['Текущий_возраст'] = df_dpo['Дата_рождения_получателя'].apply(calculate_age)
df_dpo['Возрастная_категория'] = pd.cut(df_dpo['Текущий_возраст'],[0,11,15,18,27,50,65,100],
                                        labels=['Младший возраст','12-15 лет','16-18 лет','19-27 лет','28-50 лет','51-65 лет','66 и больше'])

df_po['Текущий_возраст'] = df_po['Дата_рождения_получателя'].apply(calculate_age)
df_po['Возрастная_категория'] = pd.cut(df_po['Текущий_возраст'],[0,11,15,18,27,50,65,100],
                                        labels=['Младший возраст','12-15 лет','16-18 лет','19-27 лет','28-50 лет','51-65 лет','66 и больше'])

wb = openpyxl.load_workbook(base_file)
# Записываем лист ДПО
for ir in range(0, len(df_dpo)):
    for ic in range(0, len(df_dpo.iloc[ir])):
        wb['ДПО'].cell(2 + ir, 1 + ic).value = df_dpo.iloc[ir][ic]
# Записываем лист ПО
for ir in range(0, len(df_po)):
    for ic in range(0, len(df_po.iloc[ir])):
        wb['ПО'].cell(2 + ir, 1 + ic).value = df_po.iloc[ir][ic]
wb.save('data/Общая таблица.xlsx')

# update_spreadsheet(base_file,df_dpo,1,2,'ДПО')
# update_spreadsheet('data/Общая таблица.xlsx',df_po,1,2,'ПО')






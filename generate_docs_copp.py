"""
Скрипт для генерации документов ДПО и ПО из шаблонов
"""
import pandas as pd
import numpy as np
import os
from dateutil.parser import ParserError
from docxtpl import DocxTemplate
from docxcompose.composer import Composer
from docx import Document
from docx2pdf import convert
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from jinja2 import exceptions
import time
import datetime
import warnings
from collections import Counter


warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None
import sys
import locale
import logging
import tempfile
import re

logging.basicConfig(
    level=logging.WARNING,
    filename="error.log",
    filemode='w',
    # чтобы файл лога перезаписывался  при каждом запуске.Чтобы избежать больших простыней. По умолчанию идет 'a'
    format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
    datefmt='%H:%M:%S',
)

def create_doc_convert_date(cell):
    """
    Функция для конвертации даты при создании документов
    :param cell:
    :return:
    """
    try:
        string_date = datetime.datetime.strftime(cell, '%d.%m.%Y')
        return string_date
    except ValueError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'
    except TypeError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'


def processing_date_column(df, lst_columns):
    """
    Функция для обработки столбцов с датами. конвертация в строку формата ДД.ММ.ГГГГ
    """
    # получаем первую строку
    first_row = df.iloc[0, lst_columns]

    lst_first_row = list(first_row)  # Превращаем строку в список
    lst_date_columns = []  # Создаем список куда будем сохранять колонки в которых находятся даты
    tupl_row = list(zip(lst_columns,
                        lst_first_row))  # Создаем список кортежей формата (номер колонки,значение строки в этой колонке)

    for idx, value in tupl_row:  # Перебираем кортеж
        result = check_date_columns(idx, value)  # проверяем является ли значение датой
        if result:  # если да то добавляем список порядковый номер колонки
            lst_date_columns.append(result)
        else:  # иначе проверяем следующее значение
            continue
    for i in lst_date_columns:  # Перебираем список с колонками дат, превращаем их в даты и конвертируем в нужный строковый формат
        df.iloc[:, i] = pd.to_datetime(df.iloc[:, i], errors='coerce', dayfirst=True)
        df.iloc[:, i] = df.iloc[:, i].apply(create_doc_convert_date)

def check_date_columns(i, value):
    """
    Функция для проверки типа колонки. Необходимо найти колонки с датой
    :param i:
    :param value:
    :return:
    """
    try:
        itog = pd.to_datetime(str(value), infer_datetime_format=True)
    except:
        pass
    else:
        return i

def clean_value(value):
    """
    Функция для обработки значений колонки от  пустых пробелов,нан
    :param value: значение ячейки
    :return: очищенное значение
    """
    if value is np.nan:
        return 'Не заполнено'
    str_value = str(value)
    if str_value == '':
        return 'Не заполнено'
    elif str_value ==' ':
        return 'Не заполнено'

    return str_value

def generate_docs(path_to_folder_template:str,file_data:str,path_to_end_folder:str,**kwargs):
    """
    path_to_folder_template: путь к папке с шаблонами
    file_data: путь к таблице
    path_to_end_folder: путь к конечной папке
    dct_opti: словарь с доп параметрами название программы ,даты начала и конца курсов и т.д
    """
    df = pd.read_excel(file_data,sheet_name='ДПО',dtype=str)
    df['Дата_выдачи_документа'] = pd.to_datetime(df['Дата_выдачи_документа'], errors='coerce',dayfirst=True)
    df['Дата_рождения_получателя'] = pd.to_datetime(df['Дата_рождения_получателя'], errors='coerce',dayfirst=True)
    lst_frdo_columns = ['','','','','','','','','','','','','','','',]  # Список колонок используемых в ФРДО




    # lst_xlsx = [] # список для хранения шаблонов Excel
    # # Открывае шаблон ДПО ФИС ФРДО
    # lst_df = df.values.tolist() # превращаем в список списков
    # template_dpo_fis_frdo = openpyxl.load_workbook(f'{path_to_folder_template}/ДПО/ДПО_ФИС-ФРДО.xlsx')
    # first_sheet = template_dpo_fis_frdo.sheetnames[0]
    # start_row = 2
    # for row_data in lst_df:
    #     for col, value in enumerate(row_data, 1):
    #         template_dpo_fis_frdo[first_sheet].cell(row=start_row, column=col, value=value)
    #     start_row += 1
    # for column in template_dpo_fis_frdo[first_sheet].columns:
    #     max_length = 0
    #     column_name = get_column_letter(column[0].column)
    #     for cell in column:
    #         try:
    #             if len(str(cell.value)) > max_length:
    #                 max_length = len(cell.value)
    #         except:
    #             pass
    #     adjusted_width = (max_length + 2)
    #     template_dpo_fis_frdo[first_sheet].column_dimensions[column_name].width = adjusted_width
    #
    # template_dpo_fis_frdo.save(f'{path_to_end_folder}/ДПО ФИС-ФРДО для загрузки.xlsx')










if __name__ == '__main__':
    path_folder_template_main = 'data/example/Шаблоны'
    # data_file_main = 'data/example/ДПО_Цифровые_инструменты_в_образовательной_среде_БРИЭТ_март.xlsx'
    data_file_main = 'data/example/Тестовый вариант.xlsx'
    path_end_folder_main = 'data/example/result'


    generate_docs(path_folder_template_main,data_file_main,path_end_folder_main,name_program='Аугметика',type_program='ДПО')
    print('Lindy Booth!')

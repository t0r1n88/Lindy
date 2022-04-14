import tkinter
import numpy as np
import pandas as pd
pd.options.mode.chained_assignment = None  # default='warn'
import os
from docxtpl import DocxTemplate
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from openpyxl.styles import Alignment
from openpyxl.styles import Color, PatternFill, Font, Border
from openpyxl.styles import colors
import time
import datetime
from datetime import date
from openpyxl.chart.label import DataLabelList
from openpyxl.chart import BarChart, Reference, PieChart, PieChart3D, Series

pd.options.display.max_colwidth = 100
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
import re

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and f  or PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)
def select_file_template_doc():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_doc
    name_file_template_doc = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))


def select_file_data_doc():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_data_doc
    # Получаем путь к файлу
    name_file_data_doc = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_doc():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_doc
    path_to_end_folder_doc = filedialog.askdirectory()


name_file_data_doc = 'Для автозаполнения ОП_ПК.xlsx'
name_file_template_doc = 'Автошаблон_ПК_программа.docx'
# name_file_template_doc = 'темп.docx'
path_to_end_folder_doc = 'data'

# Открываем таблицу
base_program_df = pd.read_excel(name_file_data_doc,sheet_name='1. По программе',dtype=str)
base_program_df.fillna('',inplace=True)
base_up_df = pd.read_excel(name_file_data_doc,sheet_name='2. По дисциплинам_модулям',dtype=str)
# Незаполненые ячейки заполняем пустой строкой
# Создаем специализированные датафреймы
prepod_df = base_up_df[['ФИО_преподавателя','Научная_степень_звание_должность','Сфера_пед_интересов','Опыт_стаж','Трудовая_функция','Уровень_квалификации','Полномочия','Характер_умений','Характер_знаний']]
# удаляем пустые строки
prepod_df.dropna(axis=0,how='any',inplace=True,thresh=3)
prepod_df.fillna('',inplace=True)
up_df = base_up_df[['Номер_по_УП','Вид_раздела','Наименование_раздела','Трудоемкость','Лекции_час','Практики_час','СРС_час','Трудовая_функция','Уровень_квалификации','Содержание_раздела'
    ,'Код_ОПК_ПК_по_ФГОС','Наименование_ПК_ОПК','Содержание_СРС','ДИСЦ_Технологии_обучения_1','ДИСЦ_Технологии_обучения_2','ДИСЦ_Технологии_обучения_3']]
up_df.dropna(axis=0,how='all',inplace=True)
up_df.fillna('',inplace=True)

# Конвертируем датафрейм с описанием программы в список словарей
data_program = base_program_df.to_dict('records')
context = data_program[0]
# Добавляем ключ Профессиональный стандарт
up_df['Профессиональный_стандарт'] = data_program[0]['Профессиональный_стандарт']
# Создаем датафрейм для таблицы

# Добавляем в словарь context полностью весь список словарей data ,чтобы реализовать добавление в одну таблицу данных из разных ключей
context['prepod_lst'] = prepod_df.to_dict('records')
context['up_lst'] = up_df.to_dict('records')



doc = DocxTemplate(name_file_template_doc)
# Создаем документ
doc.render(context)
# сохраняем документ
# название программы
name_program = base_program_df['Наименование_программы'].tolist()[0]
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
doc.save(
    f'{path_to_end_folder_doc}/Программа повышения квалификации {name_program} {current_time}.docx')
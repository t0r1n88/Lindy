import tkinter
import numpy as np
import pandas as pd
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
name_file_template_doc = 'Автошаблон_ПК_программа.doc'
path_to_end_folder_doc = 'data'

# Открываем таблицу
base_program_df = pd.read_excel(name_file_data_doc,sheet_name='По программе',dtype=str)
base_up_df = pd.read_excel(name_file_data_doc,sheet_name='По дисциплинам_модулям',dtype=str)
# Создаем специализированные датафреймы
prepod_df = base_up_df[['ФИО_преподавателя','Научная_степень_звание_должность','Сфера_пед_интересов','Опыт_стаж']]
up_df = base_up_df[['Номер_по_УП','Вид_раздела','Наименование_раздела','Трудоемкость','Лекции_час','Практики_час','СРС_час','Содержание_раздела','','',]]

# Конвертируем датафрейм с описанием программы в список словарей
data_program = base_program_df.to_dict('records')

context = data_program
# Добавляем в словарь context полностью весь список словарей data ,чтобы реализовать добавление в одну таблицу данных из разных ключей
# context['lst_items'] = data
doc = DocxTemplate(name_file_template_doc)
# Создаем документ
doc.render(context)
# сохраняем документ
# генерируем название
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
doc.save(
    f'{path_to_end_folder_doc}/Программа повышения квалифкации{0name_doc} {current_time}.docx')

import tkinter
import numpy as np
import sys
import pandas as pd
import os
from docxtpl import DocxTemplate
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time
import datetime
pd.options.mode.chained_assignment = None  # default='warn'
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
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


def select_file_data_obraz():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_data_obraz_program
    # Получаем путь к файлу
    name_file_data_obraz_program = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_educ_obraz():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_obraz_program
    path_to_end_folder_obraz_program = filedialog.askdirectory()

def select_file_template_educ_program():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_educ_program
    name_file_template_educ_program = filedialog.askopenfilename(
        filetypes=(('Word files', '*.docx'), ('all files', '*.*')))

def convert_date(cell):
    """
    Функция для конвертации даты в формате 1957-05-10 в формат 10.05.1957(строковый)
    """

    try:
        string_date = datetime.datetime.strftime(cell, '%d.%m.%Y')
        return string_date
    except TypeError:
        print(cell)
        messagebox.showerror('ЦОПП Бурятия', 'Проверьте правильность заполнения ячеек с датой!!!')
        quit()
    except ValueError:
        pass
        # print(cell)
        # # messagebox.showerror('ЦОПП Бурятия', 'Пустая ячейка с датой или некорректная запись!!!')
        # # quit()
name_file_data_obraz_program = 'Для автозаполнения ОП_ПК от 02_06_2022.xlsx'
name_file_template_educ_program = 'Автошаблон_ПК_программа от 02_06_2022.docx'

# name_file_template_doc = 'темп.docx'
path_to_end_folder_obraz_program = 'data'

# Открываем таблицу
base_program_df = pd.read_excel(name_file_data_obraz_program, sheet_name='1. По программе', dtype=str)
base_program_df.fillna('',inplace=True)
base_up_df = pd.read_excel(name_file_data_obraz_program, sheet_name='2. По дисциплинам_модулям', dtype=str)

base_program_df['Дата_приказа_МИНТРУДА'] = pd.to_datetime(base_program_df['Дата_приказа_МИНТРУДА'],dayfirst=True,errors='coerce')
base_program_df['Дата_приказа_МИНТРУДА'] = base_program_df['Дата_приказа_МИНТРУДА'].apply(convert_date)

# Создаем специализованный датафрейм который включает в себя категории, технологии и пр.Т.е все что включает больше одной строки
multi_line_df = base_program_df[['Категория_слушателей','Форма_обучения','Технологии_обучения','Разработчики_программы']]
# Заменяем пустые строки на Nan
multi_line_df.replace('',np.NaN,inplace=True)
# Для технологий
tech_df = multi_line_df['Технологии_обучения']
tech_df.dropna(inplace=True)
# Для категорий
cat_df = multi_line_df['Категория_слушателей']
cat_df.dropna(inplace=True)
# для разработчиков
dev_df = multi_line_df['Разработчики_программы']
dev_df.dropna(inplace=True)

# Незаполненые ячейки заполняем пустой строкой
# Создаем специализированные датафреймы
all_prepod_df = base_up_df[['ФИО_преподавателя', 'Научная_степень_звание_должность', 'Сфера_пед_интересов', 'Опыт_стаж', 'Трудовая_функция', 'Уровень_квалификации', 'Полномочия', 'Характер_умений', 'Характер_знаний']]
# удаляем пустые строки
all_prepod_df.dropna(axis=0, how='any', inplace=True, thresh=3)
all_prepod_df.fillna('', inplace=True)
# Удаляем дубликаты преподавателей, чтобы корректно заполнять таблицу преподавательского состава
unique_prepod_df = all_prepod_df.copy()
unique_prepod_df.drop_duplicates(subset=['ФИО_преподавателя'],inplace=True,ignore_index=True)

# Удаляем дубликаты уровней квалификации
level_qual_prepod = all_prepod_df.copy()
level_qual_prepod.drop_duplicates(subset=['Уровень_квалификации'],inplace=True,ignore_index=True)

# Создаем и обрабатываем датафрейм  учебной программы
up_df = base_up_df[['Наименование_раздела','Трудоемкость','Лекции_час','Практики_час','СРС_час','Трудовая_функция','Уровень_квалификации','Код_ОПК_ПК_по_ФГОС','Наименование_ПК_ОПК']]
up_df.dropna(axis=0,how='all',inplace=True)
up_df.fillna('',inplace=True)

# Создаем датафрейм учебной программы без учета строки ИТОГО для таблиц краткой аннотации,3.3
short_up_df = up_df[up_df['Наименование_раздела'] != 'ИТОГО']




# Конвертируем датафрейм с описанием программы в список словарей
data_program = base_program_df.to_dict('records')
context = data_program[0]
# Добавляем ключ Профессиональный стандарт
up_df['Профессиональный_стандарт'] = data_program[0]['Профессиональный_стандарт']
short_up_df['Профессиональный_стандарт'] = data_program[0]['Профессиональный_стандарт']
# Создаем датафрейм для таблицы

# Добавляем в словарь context полностью весь список словарей data ,чтобы реализовать добавление в одну таблицу данных из разных ключей
context['prepod_lst'] = unique_prepod_df.to_dict('records')
context['up_lst'] = up_df.to_dict('records')
context['short_up_lst'] = short_up_df.to_dict('records')
context['qual_prepod_lst'] = level_qual_prepod.to_dict('records')

# Список для технологий обучения
context['lst_tech'] = tech_df
# Список для разработчиков
context['lst_dev'] = dev_df
#Список для категорий обучения
context['lst_cat'] = cat_df

doc = DocxTemplate(name_file_template_educ_program)
# Создаем документ
doc.render(context)
# сохраняем документ
# название программы
name_program = base_program_df['Наименование_программы'].tolist()[0]
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
doc.save(
    f'{path_to_end_folder_obraz_program}/Программа повышения квалификации {name_program} {current_time}.docx')

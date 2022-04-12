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

def create_educ_program():
    """
    Функция для генерации образовательных программ
    """
    try:
        # Открываем таблицу
        base_program_df = pd.read_excel(name_file_data_obraz_program,sheet_name='1. По программе',dtype=str)
        base_program_df.fillna('',inplace=True)
        base_up_df = pd.read_excel(name_file_data_obraz_program,sheet_name='2. По дисциплинам_модулям',dtype=str)
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
    except NameError:
        messagebox.showinfo('ЦОПП Бурятия', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
    else:
        messagebox.showinfo('ЦОПП Бурятия', 'Создание образовательной программы\nЗавершено!')


if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Бурятия')
    window.geometry('700x860')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку создания документов по шаблону
    tab_create_educ_program = ttk.Frame(tab_control)
    tab_control.add(tab_create_educ_program, text='Создание образовательных программ')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_create_educ_program,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nГенерация документов по шаблону')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')

    img = PhotoImage(file=path_to_img)
    Label(tab_create_educ_program,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)


    # Создаем кнопку Выбрать файл с данными
    btn_data_data_obraz = Button(tab_create_educ_program, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                          command=select_file_data_obraz
                          )
    btn_data_data_obraz.grid(column=0, row=2, padx=10, pady=10)

    #Создаем кнопку выбора шаблона
    # Создаем кнопку Выбрать файл с данными
    btn_template_educ_program = Button(tab_create_educ_program, text='2) Выберите шаблон', font=('Arial Bold', 20),
                          command=select_file_template_educ_program
                          )
    btn_template_educ_program.grid(column=0, row=3, padx=10, pady=10)



    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_educ_program = Button(tab_create_educ_program, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder_educ_obraz
                                       )
    btn_choose_end_folder_educ_program.grid(column=0, row=4, padx=10, pady=10)

    btn_create_educ_program = Button(tab_create_educ_program, text='4) Создать образовательную\n программу', font=('Arial Bold', 20),
                                       command=create_educ_program
                                       )
    btn_create_educ_program.grid(column=0, row=5, padx=10, pady=10)

    window.mainloop()
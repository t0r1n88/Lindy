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

# Вспомогательные для ПО
def select_file_data_obraz_po():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_data_obraz_program_po
    # Получаем путь к файлу
    name_file_data_obraz_program_po = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_end_folder_educ_obraz_po():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_obraz_program_po
    path_to_end_folder_obraz_program_po = filedialog.askdirectory()

def select_file_template_educ_program_po():
    """
    Функция для выбора файла шаблона
    :return: Путь к файлу шаблона
    """
    global name_file_template_educ_program_po
    name_file_template_educ_program_po = filedialog.askopenfilename(
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
        messagebox.showerror('Андраста ver 1.81', 'Проверьте правильность заполнения ячеек с датой!!!')
        quit()
    except ValueError:
        pass
        # print(cell)
        # # messagebox.showerror('Андраста ver 1.81', 'Пустая ячейка с датой или некорректная запись!!!')
        # # quit()

def create_educ_program():
    """
    Функция для генерации образовательных программ
    """
    try:
        # Открываем таблицу
        base_program_df = pd.read_excel(name_file_data_obraz_program, sheet_name='1. По программе', dtype=str)
        base_program_df.fillna('', inplace=True)
        # Убираем пробельные символы сначала и в конце каждой ячейки
        base_program_df = base_program_df.applymap(str.strip, na_action='ignore')
        base_up_df = pd.read_excel(name_file_data_obraz_program, sheet_name='2. По дисциплинам_модулям', dtype=str)
        base_up_df = base_up_df.applymap(str.strip, na_action='ignore')

        base_program_df['Дата_приказа_МИНТРУДА'] = pd.to_datetime(base_program_df['Дата_приказа_МИНТРУДА'],
                                                                  dayfirst=True, errors='coerce')
        base_program_df['Дата_приказа_МИНТРУДА'] = base_program_df['Дата_приказа_МИНТРУДА'].apply(convert_date)

        # Создаем специализованный датафрейм который включает в себя категории, технологии и пр.Т.е все что включает больше одной строки
        multi_line_df = base_program_df[
            ['Категория_слушателей', 'Форма_обучения', 'Технологии_обучения', 'Разработчики_программы']]
        # Заменяем пустые строки на Nan
        multi_line_df.replace('', np.NaN, inplace=True)
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
        all_prepod_df = base_up_df[
            ['ФИО_преподавателя', 'Научная_степень_звание_должность', 'Сфера_пед_интересов', 'Опыт_стаж',
             'Трудовая_функция', 'Уровень_квалификации', 'Полномочия', 'Характер_умений', 'Характер_знаний']]
        # удаляем пустые строки
        all_prepod_df.dropna(axis=0, how='any', inplace=True, thresh=3)
        all_prepod_df.fillna('', inplace=True)
        # Удаляем дубликаты преподавателей, чтобы корректно заполнять таблицу преподавательского состава
        unique_prepod_df = all_prepod_df.copy()
        unique_prepod_df.drop_duplicates(subset=['ФИО_преподавателя'], inplace=True, ignore_index=True)
        unique_prepod_df.replace('', np.NaN, inplace=True)
        unique_prepod_df.dropna(axis=0, how='any', inplace=True, subset=['ФИО_преподавателя'])

        # Удаляем дубликаты уровней квалификации
        level_qual_prepod = all_prepod_df.copy()

        level_qual_prepod.drop_duplicates(subset=['Уровень_квалификации'], inplace=True, ignore_index=True)

        # Создаем и обрабатываем датафрейм  учебной программы
        up_df = base_up_df[
            ['Наименование_раздела', 'Трудоемкость', 'Лекции_час', 'Практики_час', 'СРС_час', 'Трудовая_функция',
             'Уровень_квалификации', 'Код_ОПК_ПК_по_ФГОС', 'Наименование_ПК_ОПК']]
        up_df.dropna(axis=0, how='all', inplace=True)
        up_df.fillna('-', inplace=True)

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
        # Список для категорий обучения
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
    except IndexError:
        messagebox.showerror('Андраста ver 1.81', 'Заполните полностью строку 2 на листе 1.По программе!!!')
    except NameError:
        messagebox.showinfo('Андраста ver 1.81', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
    except FileNotFoundError:
        # сообщение на случай если путь до папки куда сохраняется файл слишком длинный
        messagebox.showerror('Андраста ver 1.81', f'Слишком длинный путь до сохраняемого файла!\nВыберите другую папку')
    except KeyError as e:
        messagebox.showerror('Андраста ver 1.81', f'Не найдено название колонки {e.args}')
    else:
        messagebox.showinfo('Андраста ver 1.81', 'Создание образовательной программы\nЗавершено!')

def create_educ_program_po():
    """
    Функция для генерации программ ПО
    """
    try:

        # Открываем таблицу
        base_program_df = pd.read_excel(name_file_data_obraz_program_po, sheet_name='1. По программе', dtype=str)
        base_program_df.fillna('', inplace=True)

        # Убираем пробельные символы сначала и в конце каждой ячейки
        base_program_df = base_program_df.applymap(str.strip, na_action='ignore')

        # Обрабатываем колнку дата приказа Минтруда
        base_program_df['Дата_приказа_МИНТРУДА'] = pd.to_datetime(base_program_df['Дата_приказа_МИНТРУДА'],
                                                                  dayfirst=True, errors='coerce')
        base_program_df['Дата_приказа_МИНТРУДА'] = base_program_df['Дата_приказа_МИНТРУДА'].apply(convert_date)

        # Создаем специализованный датафрейм который включает в себя категории, технологии и пр.Т.е все что включает больше одной строки
        multi_line_df = base_program_df[
            ['Форма_обучения', 'Уровни_квалификации', 'Технологии_обучения',
             'Разработчики_программы']]
        # Заменяем пустые строки на Nan
        multi_line_df.replace('', np.NaN, inplace=True)
        # Для технологий
        tech_df = multi_line_df['Технологии_обучения']
        tech_df.dropna(inplace=True)

        # Обрабатываем уровни квалификации чтобы превратить в строку
        # Создаем список, удаляя наны
        level_cat_df = multi_line_df['Уровни_квалификации'].dropna().to_list()
        # Создаем строку
        level_cat_str = ','.join(level_cat_df)
        # для разработчиков
        dev_df = multi_line_df['Разработчики_программы']
        dev_df.dropna(inplace=True)

        # Создаем базовый датафрейм по дисциплинам и модулям
        base_up_df = pd.read_excel(name_file_data_obraz_program_po, sheet_name='2. По дисциплинам_модулям', dtype=str)
        base_up_df = base_up_df.applymap(str.strip, na_action='ignore')
        # Незаполненые ячейки заполняем пустой строкой

        # Создаем специализированные датафреймы
        all_prepod_df = base_up_df[
            ['ФИО_преподавателя', 'Научная_степень_звание_должность', 'Сфера_пед_интересов', 'Опыт_стаж',
             'Форма_контроля', 'Уровень_квалификации', 'Полномочия', 'Характер_умений', 'Характер_знаний']]
        # удаляем пустые строки
        all_prepod_df.dropna(axis=0, how='any', inplace=True, thresh=3)
        all_prepod_df.fillna('', inplace=True)
        # Удаляем дубликаты преподавателей, чтобы корректно заполнять таблицу преподавательского состава
        unique_prepod_df = all_prepod_df.copy()
        unique_prepod_df.drop_duplicates(subset=['ФИО_преподавателя'], inplace=True, ignore_index=True)

        # Удаляем дубликаты уровней квалификации
        level_qual_prepod = all_prepod_df.copy()
        level_qual_prepod.drop_duplicates(subset=['Уровень_квалификации'], inplace=True, ignore_index=True)

        # Создаем и обрабатываем датафрейм  учебной программы
        up_df = base_up_df[
            ['Наименование_раздела', 'Трудоемкость', 'Лекции_час', 'Практики_час', 'СРС_час', 'Форма_контроля',
             'Уровень_квалификации']]
        up_df.dropna(axis=0, how='all', inplace=True)
        up_df.fillna('-', inplace=True)

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
        # Список лиц осваивающих программы
        context['lst_multi_category'] = multi_line_df.to_dict('records')
        # Список для технологий обучения
        context['lst_tech'] = tech_df

        # Список для разработчиков
        context['lst_dev'] = dev_df
        # Добавляем в контекст строку для уровней
        context['level_cat_str'] = level_cat_str

        doc = DocxTemplate(name_file_template_educ_program_po)
        # Создаем документ
        doc.render(context)
        # сохраняем документ
        # название программы
        name_program = base_program_df['Наименование_профессии'].tolist()[0]
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        doc.save(
            f'{path_to_end_folder_obraz_program_po}/Программа профессионального обучения {name_program} {current_time}.docx')

    except IndexError:
        messagebox.showerror('Андраста ver 1.81', 'Заполните полностью строку 2 на листе 1.По программе!!!')
    except NameError:
        messagebox.showinfo('Андраста ver 1.81', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
    except FileNotFoundError:
        # сообщение на случай если путь до папки куда сохраняется файл слишком длинный
        messagebox.showerror('Андраста ver 1.81', f'Слишком длинный путь до сохраняемого файла!\nВыберите другую папку')
    except KeyError as e:
        messagebox.showerror('Андраста ver 1.81', f'Не найдено название колонки {e.args}')
    else:
        messagebox.showinfo('Андраста ver 1.81', 'Создание образовательной программы\nЗавершено!')

if __name__ == '__main__':
    window = Tk()
    window.title('Андраста ver 1.81')
    window.geometry('700x600')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку создания программ повышения квалификации ПК по шаблону
    tab_create_educ_program = ttk.Frame(tab_control)
    tab_control.add(tab_create_educ_program, text='Создание программ ПК')
    tab_control.pack(expand=1, fill='both')
    #
    # Создаем вкладку создания программ профессионального обучения ПО по шаблону
    tab_create_educ_program_po = ttk.Frame(tab_control)
    tab_control.add(tab_create_educ_program_po, text='Создание программ ПО')
    tab_control.pack(expand=1, fill='both')
    #
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_create_educ_program,
                      text='Центр опережающей профессиональной подготовки\n Республики Бурятия\nГенерация программ\nповышения квалификации',font=15)
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

    btn_create_educ_program = Button(tab_create_educ_program, text='4) Создать программу ПК', font=('Arial Bold', 20),
                                       command=create_educ_program
                                       )
    btn_create_educ_program.grid(column=0, row=5, padx=10, pady=10)


    # Добавляем виджеты на вклдаку создания программ ПО

    # Создаем метку для описания назначения программы
    lbl_hello_po = Label(tab_create_educ_program_po,
                      text='Центр опережающей профессиональной подготовки\n Республики Бурятия\nГенерация программ\nпрофессионального обучения',font=15)
    lbl_hello_po.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img_po = resource_path('logo.png')

    img_po = PhotoImage(file=path_to_img_po)
    Label(tab_create_educ_program_po,
          image=img_po
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_data_data_obraz_po = Button(tab_create_educ_program_po, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                                 command=select_file_data_obraz_po
                                 )
    btn_data_data_obraz_po.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку выбора шаблона
    # Создаем кнопку Выбрать файл с данными
    btn_template_educ_program_po = Button(tab_create_educ_program_po, text='2) Выберите шаблон', font=('Arial Bold', 20),
                                       command=select_file_template_educ_program_po
                                       )
    btn_template_educ_program_po.grid(column=0, row=3, padx=10, pady=10)

    #Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_educ_program_po = Button(tab_create_educ_program_po, text='3) Выберите конечную папку',
                                                font=('Arial Bold', 20),
                                                command=select_end_folder_educ_obraz_po
                                                )
    btn_choose_end_folder_educ_program_po.grid(column=0, row=4, padx=10, pady=10)

    btn_create_educ_program_po = Button(tab_create_educ_program_po, text='4) Создать программу ПО', font=('Arial Bold', 20),
                                     command=create_educ_program_po
                                     )
    btn_create_educ_program_po.grid(column=0, row=5, padx=10, pady=10)

    window.mainloop()

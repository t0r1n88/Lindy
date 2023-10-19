"""
Графический интерфейс для программы по  генерации документов ДПО и ПО
"""
from generate_docs_copp import generate_docs # импортируем функцию генерации документов
from preparation_list import prepare_list # импортируем функцию подготовки данных списка

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import os
import sys
import logging


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and f  or PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

"""
Функции для создания контекстного меню(Копировать,вставить,вырезать)
"""


def make_textmenu(root):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # эта штука делает меню
    global the_menu
    the_menu = Menu(root, tearoff=0)
    the_menu.add_command(label="Вырезать")
    the_menu.add_command(label="Копировать")
    the_menu.add_command(label="Вставить")
    the_menu.add_separator()
    the_menu.add_command(label="Выбрать все")


def callback_select_all(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # select text after 50ms
    window.after(50, lambda: event.widget.select_range(0, 'end'))


def show_textmenu(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    e_widget = event.widget
    the_menu.entryconfigure("Вырезать", command=lambda: e_widget.event_generate("<<Cut>>"))
    the_menu.entryconfigure("Копировать", command=lambda: e_widget.event_generate("<<Copy>>"))
    the_menu.entryconfigure("Вставить", command=lambda: e_widget.event_generate("<<Paste>>"))
    the_menu.entryconfigure("Выбрать все", command=lambda: e_widget.select_range(0, 'end'))
    the_menu.tk.call("tk_popup", the_menu, event.x_root, event.y_root)



"""
Функции для подготовки
"""
def select_folder_template():
    """
    Функция для выбора папки где лежат шаблоны документов
    :return:
    """
    global glob_path_to_folder_template
    glob_path_to_folder_template = filedialog.askdirectory()

def select_data_file():
    """
    Функция для выбора исходной таблицы
    :return:
    """
    global glob_data_file
    # Получаем путь к файлу
    glob_data_file = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_end_folder():
    """
    Функция для выбора папки куда будут сохранены данные
    :return:
    """
    global glob_path_to_end_folder
    glob_path_to_end_folder = filedialog.askdirectory()


def processing_generate_docs():
    """
    Функция для генерации документов
    """
    dct_params = {} # словарь для дополнительных параметров
    try:
        # name_course= str(entry_name_course.get()) # получаем название курса
        # begin_course = str(entry_begin_course.get()) # получаем дату начала
        # end_course = str(entry_end_course.get())  # получаем дату окончания

        type_course = group_rb_type_course.get() # получаем значения тип курса ДПО или ПО

        # создаем документы
        generate_docs(glob_path_to_folder_template,glob_data_file,glob_path_to_end_folder,type_course)

    except NameError:
        messagebox.showerror('Веста Обработка таблиц и создание документов',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')


"""
Функции для вкладки подготовка файлов
"""
def select_prep_file():
    """
    Функция для выбора файла который нужно преобразовать
    :return:
    """
    global glob_prep_file
    # Получаем путь к файлу
    glob_prep_file = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_prep():
    """
    Функция для выбора папки куда будет сохранен преобразованный файл
    :return:
    """
    global glob_path_to_end_folder_prep
    glob_path_to_end_folder_prep = filedialog.askdirectory()


def processing_preparation_file():
    """
    Функция для генерации документов
    """
    try:
        prepare_list(glob_prep_file,glob_path_to_end_folder_prep)


    except NameError:
        messagebox.showerror('Линди Создание документов и отчетов ЦОПП версия 2.0',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
        logging.exception('AN ERROR HAS OCCURRED')



if __name__ == '__main__':
    window = Tk()
    window.title('Линди Создание документов и отчетов ЦОПП версия 2.0')
    window.geometry('850x970')
    window.resizable(False, False)
    # Добавляем контекстное меню в поля ввода
    make_textmenu(window)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    """
    Создаем вкладку для предварительной обработки списка
    """
    # Создаем вкладку создания документов по шаблону
    tab_preparation= ttk.Frame(tab_control)
    tab_control.add(tab_preparation, text='Подготовка файла')
    tab_control.pack(expand=1, fill='both')

    # размещаем виджеты на вкладке Подготовка файла
    lbl_hello = Label(tab_preparation,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Очистка от некорректных данных, поиск пропущенных значений,\n преобразование СНИЛС в формат ХХХ-ХХХ-ХХХ ХХ.')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка . Пришлось переименовывать переменную, иначе картинка не отображалась
    path_to_img_prep = resource_path('logo.png')
    img_prep = PhotoImage(file=path_to_img_prep)
    Label(tab_preparation,
          image=img_prep
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_prep = LabelFrame(tab_preparation, text='Подготовка')
    frame_data_prep.grid(column=0, row=1, padx=10)

    # Создаем кнопку выбора файла с данными
    btn_choose_prep_file= Button(frame_data_prep, text='1) Выберите файл', font=('Arial Bold', 20),
                                       command=select_prep_file)
    btn_choose_prep_file.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder_prep= Button(frame_data_prep, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder_prep)
    btn_choose_end_folder_prep.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку очистки
    btn_choose_processing_prep= Button(frame_data_prep, text='3) Выполнить подготовку', font=('Arial Bold', 20),
                                       command=processing_preparation_file)
    btn_choose_processing_prep.grid(column=0, row=4, padx=10, pady=10)



    """
    Создаем вкладку создания документов по шаблону
    """
    tab_create_doc = ttk.Frame(tab_control)
    tab_control.add(tab_create_doc, text='Создание документов')
    tab_control.pack(expand=1, fill='both')

    # размещаем виджеты на вкладке Прочее
    lbl_hello = Label(tab_create_doc,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nГенерация документов')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка . Пришлось переименовывать переменную, иначе картинка не отображалась
    path_to_img_doc = resource_path('logo.png')
    img_doc = PhotoImage(file=path_to_img_doc)
    Label(tab_create_doc,
          image=img_doc
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_doc = LabelFrame(tab_create_doc, text='Подготовка')
    frame_data_doc.grid(column=0, row=1, padx=10)


    # Создаем кнопку выбора папки с шаблонами
    btn_choose_folder_template = Button(frame_data_doc, text='1) Выберите папку с шаблонами', font=('Arial Bold', 20),
                                       command=select_folder_template)
    btn_choose_folder_template.grid(column=0, row=1, padx=10, pady=10)

    # Создаем кнопку выбора файла с данными
    btn_choose_data_file= Button(frame_data_doc, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                                       command=select_data_file)
    btn_choose_data_file.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder= Button(frame_data_doc, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder)
    btn_choose_end_folder.grid(column=0, row=3, padx=10, pady=10)

    # Переключатель:вариант работы
    # Создаем переключатель
    group_rb_type_course = StringVar()
    group_rb_type_course.set('ДПО') # значение по умолчанию
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type_course = LabelFrame(frame_data_doc, text='4) Выберите тип курса')
    frame_rb_type_course.grid(column=0, row=4, padx=10)
    #
    Radiobutton(frame_rb_type_course, text='ДПО', variable=group_rb_type_course,
                value='ДПО').pack()
    Radiobutton(frame_rb_type_course, text='ПО', variable=group_rb_type_course,
                value='ПО').pack()

    # frame_dop_data = LabelFrame(tab_create_doc,text='Введите дополнительные данные')
    # frame_dop_data.grid(column=0, row=5, padx=10)
    #
    #
    # # Создаем переменную для хранения названия программы
    # name_course = StringVar()
    # # пояснение
    # label_name_course = Label(frame_dop_data,text='5) Введите название курса')
    # label_name_course.grid(column=0,row=6,padx=2)
    # # поле ввода
    # entry_name_course = Entry(frame_dop_data,textvariable=name_course,width=70)
    # entry_name_course.grid(column=0,row=7,padx=10)
    # # пояснение
    # label_date_course = Label(frame_dop_data,text='6) Введите даты начала и дату окончания курса в формате ДД.ММ.ГГГГ\n'
    #                                                'Например 12.05.2023')
    # label_date_course.grid(column=0,row=8,padx=2)
    #
    # # метки для полей ввода дат
    # label_begin_course = Label(frame_dop_data,text='Дата начала курса')
    # label_begin_course.grid(column=0,row=9,sticky='w',padx=2)
    #
    # label_end_course = Label(frame_dop_data, text='Дата завершения курса')
    # label_end_course.grid(column=1, row=9, padx=2)
    #
    #
    # date_begin_course = StringVar() # переменная для начала курса
    # date_end_course = StringVar() # переменная для конца курса
    #
    # entry_begin_course = Entry(frame_dop_data,textvariable=date_begin_course,width=15)
    # entry_begin_course.grid(column=0,sticky='w',row=10,padx=0)
    #
    # entry_end_course = Entry(frame_dop_data, textvariable=date_end_course, width=15)
    # entry_end_course.grid(column=1, row=10, padx=0)











    # Создаем кнопку генерации документов
    btn_choose_processing_doc= Button(tab_create_doc, text='7) Создать документы', font=('Arial Bold', 20),
                                       command=processing_generate_docs)
    btn_choose_processing_doc.grid(column=0, row=15, padx=10, pady=10)










    window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    window.bind_class("Entry", "<Control-a>", callback_select_all)
    window.mainloop()







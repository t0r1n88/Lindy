"""
Графический интерфейс для программы по  генерации документов ДПО и ПО
"""

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import os
import sys


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and f  or PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

"""
Функции для подготовки
"""
def select_folder_template():
    """
    Функция для выбора папки где лежат шаблоны документов
    :return:
    """
    global glob_path_to_folder_template
    path_to_folder_template = filedialog.askdirectory()

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


def processing_generate_docs(path_to_folder_template:str,data_table:str,path_to_end_folder:str):
    """
    path_to_folder_template: путь к папке с шаблонами
    data_file: путь к таблице
    path_to_end_folder: путь к конечной папке
    """
    pass



if __name__ == '__main__':
    window = Tk()
    window.title('Линид Создание документов и отчетов ЦОПП версия 2.0')
    window.geometry('700x970')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку создания документов по шаблону
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
    btn_choose_data_file= Button(frame_data_doc, text='2) Выберите файл', font=('Arial Bold', 20),
                                       command=select_data_file)
    btn_choose_data_file.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку выбора конечной папки
    btn_choose_end_folder= Button(frame_data_doc, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder)
    btn_choose_end_folder.grid(column=0, row=3, padx=10, pady=10)

    # Переключатель:вариант работы
    # Создаем переключатель
    group_rb_type_course = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type_course = LabelFrame(frame_data_doc, text='4) Выберите тип курса')
    frame_rb_type_course.grid(column=0, row=4, padx=10)
    #
    Radiobutton(frame_rb_type_course, text='ДПО', variable=group_rb_type_course,
                value=0).pack()
    Radiobutton(frame_rb_type_course, text='ПО', variable=group_rb_type_course,
                value=1).pack()

    frame_dop_data = LabelFrame(tab_create_doc,text='Введите дополнительные данные')
    frame_dop_data.grid(column=0, row=5, padx=10)


    # Создаем переменную для хранения названия программы
    name_course = StringVar()
    # пояснение
    label_name_course = Label(frame_dop_data,text='5) Введите название курса')
    label_name_course.grid(column=0,row=6,padx=2)
    # поле ввода
    entry_name_course = Entry(frame_dop_data,textvariable=name_course,width=50)
    entry_name_course.grid(column=0,row=7,padx=10)
    # пояснение
    label_date_course = Label(frame_dop_data,text='6) Введите дату начала и дату окончания курса в формате ДД.ММ.ГГГГ\n'
                                                   'Например 12.05.2023')
    label_date_course.grid(column=0,row=8,padx=2)

    # метки для полей ввода дат
    label_begin_course = Label(frame_dop_data,text='Дата начала курса')
    label_begin_course.grid(column=0,row=9,sticky='w',padx=2)

    label_end_course = Label(frame_dop_data, text='Дата завершения курса')
    label_end_course.grid(column=1, row=9, padx=2)


    date_begin_course = StringVar() # переменная для начала курса
    date_end_course = StringVar() # переменная для конца курса

    entry_begin_course = Entry(frame_dop_data,textvariable=date_begin_course,width=15)
    entry_begin_course.grid(column=0,sticky='w',row=10,padx=0)

    entry_end_course = Entry(frame_dop_data, textvariable=date_end_course, width=15)
    entry_end_course.grid(column=1, row=10, padx=0)











    # Создаем кнопку генерации документов
    btn_choose_processing_doc= Button(tab_create_doc, text='7) Создать документы', font=('Arial Bold', 20),
                                       command=processing_generate_docs)
    btn_choose_processing_doc.grid(column=0, row=15, padx=10, pady=10)












    window.mainloop()







import pandas as pd
import os
from docxtpl import DocxTemplate
import csv
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk


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


def generate_docs_dpo():
    """
    Функция для создания ддокументов по ДПО
    :return:
    """
    try:
        # Считываем данные с листа ДПО в указанной таблице
        df = pd.read_excel(name_file_data_doc, sheet_name='ДПО')

        # Конвертируем датафрейм в список словарей
        data = df.to_dict('records')

        # Создаем переменную для типа создаваемого документа
        status_rb_type_doc = group_rb_type_doc.get()
        # если статус == 0 то создаем индивидуальные приказы по количеству строк.30 строк-30 документов
        if status_rb_type_doc == 0:
            print('Индивидуальный')
            try:
                for row in data:
                    doc = DocxTemplate(name_file_template_doc)
                    context = row
                    # Превращаем строку в список кортежей, где первый элемент кортежа это ключ а второй данные

                    doc.render(context)

                    doc.save(f'{path_to_end_folder_doc}/{row["ФИО_именительный"]}.docx')
            except:
                messagebox.showerror('ЦОПП Бурятия','Проверьте содержимое шаблона\nНе допускаются любые символы кроме _ в словах внутри фигурных скобок\nСлова должны могут быть разделены нижним подчеркиванием')
                exit()

            else:
                messagebox.showinfo('ЦОПП Бурятия', 'Создание документов успешно завершено!')

        else:
            print('Групповой')

            # Создаем список в котором будет храниьт ФИО
            lst_students = []

            # Итеруемся по списку словарей, чтобы получить список ФИО
            try:
                for row in data:
                    lst_students.append(row['ФИО_именительный'])
                # Получаем первую строку таблицы, предполагая что раз это групповой список то и данные будут совпадать
                context = data[0]
                # Создаем в context  пару ключ:значение lst_studenst:список студентов
                context['lst_students'] = lst_students
                # Загружаем шаблон
                doc = DocxTemplate(name_file_template_doc)

                # Создаем документ
                doc.render(context)
                # сохраняем документ
                doc.save(f'{path_to_end_folder_doc}/{context["Порядковый_номер_группы"]}.docx')
            except KeyError:
                messagebox.showerror('ЦОПП Бурятия,Колонка с ФИО должна называться ФИО_именительный')
                exit()

            except OSError:
                messagebox.showerror('ЦОПП Бурятия','Закройте открытый файл Word')
                exit()
            except:
                messagebox.showerror('ЦОПП Бурятия',
                                     'Проверьте содержимое шаблона\nНе допускаются любые символы кроме _ в словах внутри фигурных скобок\nСлова должны могут быть разделены нижним подчеркиванием')
                exit()
            else:
                messagebox.showinfo('ЦОПП Бурятия', 'Создание документов успешно завершено!')

    except NameError as e:
        messagebox.showinfo('ЦОПП Бурятия', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')


def generate_docs_po():
    """
    Функция для создания документов ПО
    :return:
    """
    try:
        # Считываем данные с листа ДПО в указанной таблице
        df = pd.read_excel(name_file_data_doc, sheet_name='ПО')

        # Конвертируем датафрейм в список словарей
        data = df.to_dict('records')

        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_doc)
            context = row
            # Превращаем строку в список кортежей, где первый элемент кортежа это ключ а второй данные
            id_row = list(row.items())
            try:
                doc.render(context)

                doc.save(f'{path_to_end_folder_doc}/{id_row[0][1]}.docx')
            except:
                messagebox.showerror('ЦОПП Бурятия',
                                     'Проверьте содержимое шаблона\nНе допускаются любые символы кроме _ в словах внутри фигурных скобок\nСлова должны могут быть разделены нижним подчеркиванием')
                continue
        messagebox.showinfo('ЦОПП Бурятия', 'Создание документов успешно завершено!')
    except NameError as e:
        messagebox.showinfo('ЦОПП Бурятия', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')


def generate_docs_other():
    """
    Функция для создания документов из произвольных таблиц(т.е. отличающихся от структуры базы данных ЦОПП Бурятия)
    :return:
    """
    try:
        # Считываем данные
        df = pd.read_excel(name_file_data_doc)

        # Конвертируем датафрейм в список словарей
        data = df.to_dict('records')

        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_doc)
            context = row
            # Превращаем строку в список кортежей, где первый элемент кортежа это ключ а второй данные
            id_row = list(row.items())
            try:
                doc.render(context)

                doc.save(f'{path_to_end_folder_doc}/{id_row[0][1]}.docx')
            except:
                messagebox.showerror('ЦОПП Бурятия',
                                     'Проверьте содержимое шаблона\nНе допускаются любые символы кроме _ в словах внутри фигурных скобок\nСлова должны могут быть разделены нижним подчеркиванием')
                continue
        messagebox.showinfo('ЦОПП Бурятия', 'Создание документов успешно завершено!')
    except NameError as e:
        messagebox.showinfo('ЦОПП Бурятия', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')


if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Бурятия')
    window.geometry('650x860')
    window.resizable(False, False)

    # path_to_icon = resource_path('favicon.ico')
    # window.iconbitmap(path_to_icon)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку создания документов по шаблону
    tab_create_doc = ttk.Frame(tab_control)
    tab_control.add(tab_create_doc, text='Создание документов')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_create_doc,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nГенерация документов по шаблону')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Переключатель:индивидуальный или списочный приказл
    # Создаем переменную хранящую тип документа, в зависимости от значения будет использоваться та или иная функция
    group_rb_type_doc = IntVar()

    frame_rb_type_doc = LabelFrame(tab_create_doc, text='Выберите тип создаваемого документа')
    frame_rb_type_doc.grid(column=0, row=1, padx=10)

    Radiobutton(frame_rb_type_doc, text='Индивидуальные документы', variable=group_rb_type_doc, value=0).pack()
    Radiobutton(frame_rb_type_doc, text='Списочный документ', variable=group_rb_type_doc, value=1).pack()

    # Картинка
    path_to_img = resource_path('logo.png')
    img = PhotoImage(file=path_to_img)
    Label(tab_create_doc,
          image=img
          ).grid(column=0, row=2, padx=10, pady=25)

    # Создаем кнопку Выбрать шаблон
    btn_template_contract = Button(tab_create_doc, text='1) Выберите шаблон документа', font=('Arial Bold', 20),
                                   command=select_file_template_doc
                                   )
    btn_template_contract.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку Выбрать файл с данными
    btn_data_contract = Button(tab_create_doc, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                               command=select_file_data_doc
                               )
    btn_data_contract.grid(column=0, row=4, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_contract = Button(tab_create_doc, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                            command=select_end_folder_doc
                                            )
    btn_choose_end_folder_contract.grid(column=0, row=5, padx=10, pady=10)

    # Создаем кнопку для запуска функции генерации файлов ДПО

    btn_create_files_dpo = Button(tab_create_doc, text='Создать документы ДПО', font=('Arial Bold', 20),
                                  command=generate_docs_dpo
                                  )
    btn_create_files_dpo.grid(column=0, row=6, padx=10, pady=10)

    # Создаем кнопку для запуска функции генерации файлов ПО
    btn_create_files_po = Button(tab_create_doc, text='Создать документы ПО', font=('Arial Bold', 20),
                                 command=generate_docs_po
                                 )
    btn_create_files_po.grid(column=0, row=7, padx=10, pady=10)

    # Создаем кнопку для создания документов из таблиц с произвольной структурой
    btn_create_files_other = Button(tab_create_doc, text='Создать документы из произвольной таблицы',
                                    font=('Arial Bold', 20),
                                    command=generate_docs_other
                                    )
    btn_create_files_other.grid(column=0, row=8, padx=10, pady=10)

    window.mainloop()
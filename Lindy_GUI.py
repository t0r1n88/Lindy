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


def select_file_data_report():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global name_file_data_report
    # Получаем путь к файлу
    name_file_data_report = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_report():
    """
    Функция для выбора папки куда будут генерироваться файлы
    :return:
    """
    global path_to_end_folder_report
    path_to_end_folder_report = filedialog.askdirectory()



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
            except KeyError:
                messagebox.showerror('ЦОПП Бурятия','Колонка с ФИО должна называться ФИО_именительный')
            except:
                messagebox.showerror('ЦОПП Бурятия','Проверьте содержимое шаблона\nНе допускаются любые символы кроме _ в словах внутри фигурных скобок\nСлова должны могут быть разделены нижним подчеркиванием')
                exit()

            else:
                messagebox.showinfo('ЦОПП Бурятия', 'Создание документов успешно завершено!')

        else:
            print('Групповой')

            # Создаем список в котором будет хранить ФИО
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

        count = 0
        # Создаем в цикле документы
        for row in data:
            doc = DocxTemplate(name_file_template_doc)
            context = row
            count += 1
            # Превращаем строку в список кортежей, где первый элемент кортежа это ключ а второй данные

            try:
                if 'ФИО' in row:
                    doc.render(context)

                    doc.save(f'{path_to_end_folder_doc}/{row["ФИО"]}.docx')
                else:
                    doc.render(context)

                    doc.save(f'{path_to_end_folder_doc}/{count}.docx')


            except:
                messagebox.showerror('ЦОПП Бурятия',
                                     'Проверьте содержимое шаблона\nНе допускаются любые символы кроме _ в словах внутри фигурных скобок\nСлова должны могут быть разделены нижним подчеркиванием')
                exit()
        messagebox.showinfo('ЦОПП Бурятия', 'Создание документов успешно завершено!')
    except NameError as e:
        messagebox.showinfo('ЦОПП Бурятия', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')

# Функции для создания отчетов
def create_report_one_pk():
    """
    Функция для создания отчета 1-ПК
    :return:
    """
    pass

def create_report_svod():
    """
    Функция для создания отчета по сводным показателям ЦОПП
    :return:
    """
    pass






if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Бурятия')
    window.geometry('700x860')
    window.resizable(False, False)

    # path_to_icon = resource_path('favicon.ico')
    # window.iconbitmap(path_to_icon)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку создания документов по шаблону
    tab_create_doc = ttk.Frame(tab_control)
    tab_control.add(tab_create_doc, text='Создание документов')
    tab_control.pack(expand=1, fill='both')

    # Создаем вкладку создания отчетов
    tab_create_report = ttk.Frame(tab_control)
    tab_control.add(tab_create_report,text='Создание отчетов')
    tab_control.pack(expand=1,fill='both')

    # Создаем вкладку для Прочих операций
    tab_create_other = ttk.Frame(tab_control)
    tab_control.add(tab_create_other,text='Прочие операции')
    tab_control.pack(expand=1,fill='both')





    # Добавляем виджеты на вкладку Создание документов
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_create_doc,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nГенерация документов по шаблону')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')
    img = PhotoImage(file=path_to_img)
    Label(tab_create_doc,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Переключатель:индивидуальный или списочный приказл
    # Создаем переменную хранящую тип документа, в зависимости от значения будет использоваться та или иная функция
    group_rb_type_doc = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type_doc = LabelFrame(tab_create_doc, text='Выберите тип создаваемого документа')
    frame_rb_type_doc.grid(column=0, row=1, padx=10)
    #
    Radiobutton(frame_rb_type_doc, text='Индивидуальные документы', variable=group_rb_type_doc, value=0).pack()
    Radiobutton(frame_rb_type_doc, text='Списочный документ', variable=group_rb_type_doc, value=1).pack()

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_doc = LabelFrame(tab_create_doc, text='Подготовка')
    frame_data_for_doc.grid(column=0, row=2, padx=10)

    # Создаем кнопку Выбрать шаблон
    btn_template_doc = Button(frame_data_for_doc, text='1) Выберите шаблон документа', font=('Arial Bold', 20),
                              command=select_file_template_doc
                              )
    btn_template_doc.grid(column=0, row=3, padx=10, pady=10)
    #
    # Создаем кнопку Выбрать файл с данными
    btn_data_doc = Button(frame_data_for_doc, text='2) Выберите файл с данными', font=('Arial Bold', 20),
                          command=select_file_data_doc
                          )
    btn_data_doc.grid(column=0, row=4, padx=10, pady=10)
    #
    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_doc = Button(frame_data_for_doc, text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder_doc
                                       )
    btn_choose_end_folder_doc.grid(column=0, row=5, padx=10, pady=10)
    #
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
    btn_create_files_other = Button(tab_create_doc, text='Создать документы\n из произвольной таблицы',
                                    font=('Arial Bold', 20),
                                    command=generate_docs_other
                                    )
    btn_create_files_other.grid(column=0, row=8, padx=10, pady=10)

# Добавляем виджеты на вкладку создания отчетов
    lbl_hello = Label(tab_create_report,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nСоздание отчетов')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка . Пришлось переименовывать переменную, иначе картинка не отображалась
    path_to_img_report = resource_path('logo.png')
    img_report = PhotoImage(file=path_to_img_report)
    Label(tab_create_report,
          image=img_report
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_report = LabelFrame(tab_create_report, text='Подготовка')
    frame_data_for_report.grid(column=0, row=2, padx=10)

    # Создаем кнопку Выбрать файл с данными
    btn_data_report = Button(frame_data_for_report, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                               command=select_file_data_report
                               )
    btn_data_report.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_report = Button(frame_data_for_report, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                            command=select_end_folder_report
                                            )
    btn_choose_end_folder_report.grid(column=0, row=5, padx=10, pady=10)

    # Создаем облать для размещения кнопок создания отчетов
    frame_create_report = LabelFrame(tab_create_report, text='Создание отчетов')
    frame_create_report.grid(column=0,row=6,padx=10)

    # Создание сводного отчета по показателям ЦОПП

    btn_report_svod = Button(frame_create_report, text='Создать сводный отчет', font=('Arial Bold', 20),
                               command=create_report_svod
                               )
    btn_report_svod.grid(column=0,row=7,padx=10,pady=10)

    btn_report_one_pk = Button(frame_create_report, text='Создать отчет 1-ПК', font=('Arial Bold', 20),
                               command=create_report_one_pk
                               )
    btn_report_one_pk.grid(column=0,row=8,padx=10,pady=10)


    #размещаем виджеты на вкладке Прочее
    lbl_hello = Label(tab_create_other,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nПрочие операции')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка . Пришлось переименовывать переменную, иначе картинка не отображалась
    path_to_img_other = resource_path('logo.png')
    img_other = PhotoImage(file=path_to_img_report)
    Label(tab_create_other,
          image=img_other
          ).grid(column=1, row=0, padx=10, pady=25)







    window.mainloop()
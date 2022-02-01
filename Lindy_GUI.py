import pandas as pd
import os
from docxtpl import DocxTemplate
import csv
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import openpyxl
import time


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

def select_file_template_table():
    """
    Функция для выбора шаблона для создания общей таблицы
    :return:
    """
    global name_file_template_table
    # Получаем путь к файлу
    name_file_template_table = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))



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


def select_files_data_other():
    """
    Функция для выбора файлов с данными при выполнении прочих операций
    :return:
    """
    # Создаем глобальную переменную, дада я знаю что надо все сделать в виде классов.Потом когда нибудь
    global names_files_data_other
    names_files_data_other = filedialog.askopenfilenames(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))



def generate_docs_dpo():
    """
    Функция для создания ддокументов по ДПО
    :return:
    """
    # Считываем данные с листа ДПО в указанной таблице
    df = pd.read_excel(name_file_data_doc, sheet_name='ДПО')

    # Конвертируем датафрейм в список словарей
    data = df.to_dict('records')

    # Создаем переменную для типа создаваемого документа
    status_rb_type_doc = group_rb_type_doc.get()
    # если статус == 0 то создаем индивидуальные приказы по количеству строк.30 строк-30 документов
    if status_rb_type_doc == 0:
        for row in data:
            doc = DocxTemplate(name_file_template_doc)
            context = row
            # Превращаем строку в список кортежей, где первый элемент кортежа это ключ а второй данные

            doc.render(context)

            doc.save(f'{path_to_end_folder_doc}/{row["ФИО_именительный"]}.docx')
        messagebox.showinfo('ЦОПП Бурятия', 'Создание документов успешно завершено!')

    else:

        # Создаем список в котором будет храниьт ФИО
        lst_students = []

        # Итеруемся по списку словарей, чтобы получить список ФИО
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
        doc.save(f'{path_to_end_folder_doc}/Приказ по группе {context["Наименование_дополнительной_профессиональной_программы"]}.docx')
        messagebox.showinfo('ЦОПП Бурятия', 'Создание документов успешно завершено!')


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
                doc.save(f'{path_to_end_folder_doc}/Приказ по группе {context["Наименование_дополнительной_профессиональной_программы"]}.docx')
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

def create_general_table():
    """
    Функция для создания общей таблицы с данными всех групп из множества отдельных таблицы на каждую группу
    :return:
    """
    try:
        # Получаем базовые датафреймы
        df_dpo = pd.read_excel(name_file_template_table,sheet_name='ДПО')
        df_po = pd.read_excel(name_file_template_table,sheet_name='ПО')

        # Перебираем файлы собирая данные в промежуточные датафреймы и добавляя их в базовые
        for file in names_files_data_other:
            # Создаем промежуточный датафрейм с данными с листа ДПО
            temp_dpo = pd.read_excel(file, sheet_name='ДПО')
            # Создаем промежуточный датафрейм с данными с листа ДПО
            temp_po = pd.read_excel(file, sheet_name='ПО')
            # Добавляем промежуточные датафреймы в исходные
            df_dpo = df_dpo.append(temp_dpo, ignore_index=True)
            df_po = df_po.append(temp_po, ignore_index=True)

        # Код сохранения датафрейма в разные листы и сохранением форматирования  взят отсюда https://azzrael.ru/python-pandas-openpyxl-excel
        wb = openpyxl.load_workbook(name_file_template_table)
        # Записываем лист ДПО
        for ir in range(0, len(df_dpo)):
            for ic in range(0, len(df_dpo.iloc[ir])):
                wb['ДПО'].cell(2 + ir, 1 + ic).value = df_dpo.iloc[ir][ic]
        # Записываем лист ПО
        for ir in range(0, len(df_po)):
            for ic in range(0, len(df_po.iloc[ir])):
                wb['ПО'].cell(2 + ir, 1 + ic).value = df_po.iloc[ir][ic]
        # Получаем текущее время для того чтобы использовать в названии

        t = time.localtime()
        current_time = time.strftime('%d_%m_%y', t)
        #Сохраняем итоговый файл
        wb.save(f'{path_to_end_folder_doc}/Общая таблица слушателей ЦОПП от {current_time}.xlsx')
    except:
        messagebox.showerror('ЦОПП Бурятия','Возникла ошибка,проверьте шаблон таблицы\nДобавляемы файлы должны иметь одинаковую структуру с шаблоном таблицы')
    else:
        messagebox.showinfo('ЦОПП Бурятия','Общая таблица успешно создана!')


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

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_other = LabelFrame(tab_create_other, text='Подготовка')
    frame_data_for_other.grid(column=0, row=2, padx=10)

    # Создаем кнопку для выбора шаблона таблицы
    btn_table_other_template = Button(frame_data_for_other, text='Выберите шаблон таблицы', font=('Arial Bold', 20),
                              command=select_file_template_table
                              )
    btn_table_other_template.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку Выбрать файлы с данными
    btn_data_other = Button(frame_data_for_other, text='Выберите файлы с данными', font=('Arial Bold', 20),
                              command=select_files_data_other
                              )
    btn_data_other.grid(column=0, row=4, padx=10, pady=10)
    #
    btn_choose_end_folder_doc = Button(frame_data_for_other, text='Выберите конечную папку', font=('Arial Bold', 20),
                                       command=select_end_folder_doc
                                       )
    btn_choose_end_folder_doc.grid(column=0, row=5, padx=10, pady=10)

    # Кнопка создать общую таблицу

    btn_create_general_table = Button(tab_create_other, text='Создать общую таблицу', font=('Arial Bold', 20),
                               command=create_general_table
                               )
    btn_create_general_table.grid(column=0,row=6,padx=10,pady=10)






    window.mainloop()
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
from datetime import date
from openpyxl.chart.label import DataLabelList
from openpyxl.chart import BarChart, Reference,PieChart,PieChart3D,Series
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

def calculate_age(born):
    """
    Функция для расчета текущего возраста взято с https://stackoverflow.com/questions/2217488/age-from-birthdate-in-python/9754466#9754466
    :param born: дата рождения
    :return: возраст
    """

    try:
        today = date.today()
        return today.year - born.year - ((today.month, today.day) < (born.month, born.day))
    except:
        messagebox.showerror('ЦОПП Бурятия','Отсутствует или некорректная дата рождения слушателя\nПроверьте файл!')
        exit()


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
        # Создаем словарь для всех колонок
        # main_dict = dict.fromkeys(data[0].keys(),[])
        # print(main_dict)


        # Итеруемся по списку словарей, чтобы получить список ФИО

        for row in data:
            lst_students.append(row['ФИО_именительный'])



        # Получаем первую строку таблицы, предполагая что раз это групповой список то и данные будут совпадать
        context = data[0]
        # Создаем в context  пару ключ:значение lst_studenst:список студентов
        context['lst_students'] = lst_students
        # context['список_обучающихся'] = lst_students
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
    # Загружаем датафреймы
    try:
        dpo_df = pd.read_excel(name_file_data_report, sheet_name='ДПО')
        po_df = pd.read_excel(name_file_data_report, sheet_name='ПО')

        # Заполняем пустые поля для удобства группировки
        dpo_df = dpo_df.fillna('Не заполнено!!!')
        po_df = po_df.fillna('Не заполнено!!!')
        # Создаем переменную для хранения строки на которой заканчивается предыдущий показатель
        border_row = 2
        border_column = 2

        # Получение общего количества прошедших обучение,количества прошедших по ДПО,по ПО
        total_students, total_students_dpo, total_students_po = counting_total_student(dpo_df, po_df)

        # Количество обучившихся ДПО и ПО

        # Получение количества обучившихся по видам
        df_counting_type_and_name_trainning = counting_type_of_training(dpo_df, po_df)

        # Создаем новый excel файл
        wb = openpyxl.Workbook()

        # Получаем активный лист
        sheet = wb.active
        sheet.title = 'Сводные данные'

        # Начинаем заполнение листа
        # Заполняем количество обучившихся, общее и по типам
        sheet['A1'] = 'Наименование показателя'
        sheet['A2'] = 'Количество прошедших обучение ДПО'
        sheet['A3'] = 'Количество прошедших обучение ПО'
        sheet['A4'] = 'Общее количество прошедших обучение в ЦОПП'

        sheet['B1'] = 'Количество обучающихся'
        sheet['B2'] = total_students_dpo
        sheet['B3'] = total_students_po
        sheet['B4'] = total_students

        # Добавляем круговую диаграмму
        pie_main = PieChart()
        labels = Reference(sheet,min_col=1,min_row=2,max_row=3)
        data = Reference(sheet,min_col=2,min_row=2,max_row=3)

        # Для отображения данных на диаграмме
        series = Series(data, title='Series 1')
        pie_main.append(series)

        s1 = pie_main.series[0]
        s1.dLbls = DataLabelList()
        s1.dLbls.showVal = True

        pie_main.add_data(data,titles_from_data=True)
        pie_main.set_categories(labels)
        pie_main.title = 'Распределение обучившихся'
        sheet.add_chart(pie_main,'F2')
        # # Добавляем таблицу с по направлениям

        sheet['A7'] = 'Вид обучения'
        sheet['B7'] = 'Название программы'
        sheet['C7'] = 'Количество обучающихся'


        for row in df_counting_type_and_name_trainning.values.tolist():
            sheet.append(row)
        # Получаем последние активные ячейки чтобы записывалось по порядку и не налазило друг на друга
        min_column = wb.active.min_column
        max_column = wb.active.max_column
        min_row = wb.active.min_row
        max_row = wb.active.max_row

        sheet[f'A{max_row + 2}'] = 'Общее распределение прошедших обучение по полу'
        total_sex = counting_total_sex(dpo_df, po_df)
        # Добавляем в файл таблицу с распределением по полам
        for row in total_sex.values.tolist():
            sheet.append(row)

        # Получаем последние активные ячейки чтобы записывалось по порядку и не налазило друг на друга
        min_column = wb.active.min_column
        max_column = wb.active.max_column
        min_row = wb.active.min_row
        max_row = wb.active.max_row

        # Добавляем таблицу с разбиением по возрастам
        sheet[f'A{max_row + 2}'] = 'Общее распределение обучающихся по возрасту'
        age_distribution = counting_age_distribution(dpo_df, po_df)
        for row in age_distribution.values.tolist():
            sheet.append(row)

        #Добавляем круговую диаграмму
        pie_age = PieChart()
        # Для того чтобы не зависело от количества строк в предыдущих таблицах
        labels = Reference(sheet, min_col=1, min_row=max_row + 3, max_row=max_row + 2 + len(age_distribution))
        data = Reference(sheet, min_col=2, min_row=max_row + 3, max_row=max_row + 2 + len(age_distribution))
        # Для отображения данных на диаграмме
        series = Series(data, title='Series 1')
        pie_age.append(series)

        s1 = pie_age.series[0]
        s1.dLbls = DataLabelList()
        s1.dLbls.showVal = True

        pie_age.add_data(data, titles_from_data=True)
        pie_age.set_categories(labels)
        pie_age.title = 'Распределение обучившихся по возрастным категориям'

        sheet.add_chart(pie_age, f'F{max_row + 2}')

        min_column = wb.active.min_column
        max_column = wb.active.max_column
        min_row = wb.active.min_row
        max_row = wb.active.max_row

        sheet.column_dimensions['A'].width = 50
        sheet.column_dimensions['B'].width = 30
        # Сохраняем файл
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        wb.save(f'{path_to_end_folder_report}/Сводный отчет {current_time}.xlsx')
    except ValueError:
        messagebox.showerror('ЦОПП Бурятия', 'Проверьте названия листов! Должно быть ДПО и ПО')
    except KeyError:
        messagebox.showerror('ЦОПП Бурятия','Названия колонок не совпадают')
    except:
        messagebox.showerror('ЦОПП Бурятия',
                             'Возникла ошибка')
    else:
        messagebox.showinfo('ЦОПП Бурятия', 'Сводный отчет успешно создан!')



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
            #
            df_dpo = pd.concat([df_dpo,temp_dpo],ignore_index=True)
            df_po = pd.concat([df_po,temp_po],ignore_index=True)
            # df_po = df_po.append(temp_po, ignore_index=True)
        df_dpo['Текущий_возраст'] = df_dpo['Дата_рождения_получателя'].apply(calculate_age)
        df_dpo['Возрастная_категория'] = pd.cut(df_dpo['Текущий_возраст'], [0, 11, 15, 18, 27, 50, 65, 100],
                                                labels=['Младший возраст', '12-15 лет', '16-18 лет', '19-27 лет',
                                                        '28-50 лет', '51-65 лет', '66 и больше'])
        #
        df_po['Текущий_возраст'] = df_po['Дата_рождения_получателя'].apply(calculate_age)
        df_po['Возрастная_категория'] = pd.cut(df_po['Текущий_возраст'], [0, 11, 15, 18, 27, 50, 65, 100],
                                               labels=['Младший возраст', '12-15 лет', '16-18 лет', '19-27 лет',
                                                       '28-50 лет', '51-65 лет', '66 и больше'])

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

# Функции для создания сводной таблицы
def counting_total_student(dpo_df, po_df):
    """
    Функция для подсчета общего количества студентов обучающихся в цопп, и количества обучающихся по ДПО И ПО
    :param dpo_df: датафрейм ДПО
    :param po_df: датафрейм ПО
    :return: кортеж вида: общее количество обучающихся,количество обучающихся ДПО,количество обучающихся ПО
    """
    # количество по типам
    total_dpo = dpo_df.shape[0]
    total_po = po_df.shape[0]
    # общее количество
    total = total_dpo + total_po

    return total, total_dpo, total_po

def counting_type_of_training(dpo,po):
    """
    Функция для создания сводной таблицы по категориям направление подготовки, название программы,количество обучающихся
    :param dpo: датафрейм ДПО
    :param po: датафрейм ПО
    :return: датафрейм сводной таблицы
    """
    # Создаем сводные таблицы
    dpo_svod_category_and_name = pd.pivot_table(dpo, index=[
        'Дополнительная_профессиональная_программа_повышение_квалификации_профессиональная_переподготовка',
        'Наименование_дополнительной_профессиональной_программы'],
                                                values=['ФИО_именительный'],
                                                aggfunc='count')
    po_svod_category_and_name = pd.pivot_table(po,
                                               index=['Программа_профессионального_обучения_направление_подготовки',
                                                      'Наименование_программы_профессионального_обучения'],
                                               values=['ФИО_именительный'],
                                               aggfunc='count')

    # Добавляем цифровой индекс
    dpo_svod_category_and_name = dpo_svod_category_and_name.reset_index()
    po_svod_category_and_name = po_svod_category_and_name.reset_index()
    # Изменяем названия колонок, чтобы без проблем соединить 2 датафрейма
    dpo_svod_category_and_name.columns = ['Направление подготовки', 'Название программы', 'Количество обученных']
    po_svod_category_and_name.columns = ['Направление подготовки', 'Название программы', 'Количество обученных']
    # Создаем единую сводную таблицу
    general_svod_category_and_name = pd.concat([dpo_svod_category_and_name, po_svod_category_and_name],
                                               ignore_index=True)
    return general_svod_category_and_name

def counting_total_sex(dpo,po):
    """
    Функция для подсчета количества мужчин и женщин
    :param dpo: датафрейм ДПО
    :param po: датафрейм ПО
    :return: датафрейм сводной таблицы
    """
    # Создаем сводные таблицы
    dpo_total_sex = pd.pivot_table(dpo,index=['Пол_получателя'],
                                   values=['ФИО_именительный'],
                                   aggfunc='count')
    po_total_sex = pd.pivot_table(po,index=['Пол_получателя'],
                                  values=['ФИО_именительный'],
                                  aggfunc='count')
    # Извлекаем индексы
    dpo_total_sex = dpo_total_sex.reset_index()
    po_total_sex = po_total_sex.reset_index()
    #Переименовываем колонки
    dpo_total_sex.columns = ['Пол','Количество']
    po_total_sex.columns = ['Пол','Количество']

    # Соединяем в единую таблицу
    general_total_sex = pd.concat([dpo_total_sex,po_total_sex],ignore_index=True)
    #Группируем по полю Пол чтобы суммировать значения
    sum_general_total_sex = general_total_sex.groupby(['Пол']).sum().reset_index()
    return sum_general_total_sex

def counting_age_distribution(dpo,po):
    """
    Функция для подсчета количества обучающихся по возрастным категориям
    :param dpo: датафрейм ДПО
    :param po: датафрейм ПО
    :return: датафрейм сводной таблицы
    """
    #Создаем сводные таблицы
    dpo_age_distribution = pd.pivot_table(dpo,index=['Возрастная_категория'],
                                          values=['ФИО_именительный'],
                                          aggfunc='count')
    po_age_distribution = pd.pivot_table(po,index=['Возрастная_категория'],
                                          values=['ФИО_именительный'],
                                          aggfunc='count')
    # Извлекам индексы
    dpo_age_distribution = dpo_age_distribution.reset_index()
    po_age_distribution = po_age_distribution.reset_index()
    # Меняем колонки
    dpo_age_distribution.columns = ['Возрастная_категория','Количество']
    po_age_distribution.columns = ['Возрастная_категория','Количество']

    #Создаем единую сводную таблицу
    general_age_distribution = pd.concat([dpo_age_distribution,po_age_distribution],ignore_index=True)
    #Повторно группируем чтобы соединить категории из обеих таблиц
    general_age_distribution = general_age_distribution.groupby(['Возрастная_категория']).sum().reset_index()

    return general_age_distribution

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
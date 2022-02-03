import pandas as pd
import openpyxl
import time
from openpyxl.styles import Font

from openpyxl.chart import BarChart, Reference
from copy import deepcopy


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



# def counting_type_of_training(dpo,po):
#     """
#     Функция для подсчета количества студентов на каждом из видов обучения(повышение квалификации, переподготовка и т.д)
#     :param dpo_df: датафрейм ДПО
#     :param po_df: датафрейм ПО
#     :return: датафрейм вида  Вид обучения - количество студентов
#     Так как названия колонок в датафреймах отличаются придется обрабатывать по отдельности
#     """
#     # Обрабатываем ДПО
#
#     group_dpo = dpo[
#         'Дополнительная_профессиональная_программа_повышение_квалификации_профессиональная_переподготовка'].value_counts()
#
#     # Обрабатываем ПО
#
#     group_po = po['Программа_профессионального_обучения_направление_подготовки'].value_counts()
#
#     # Соединяем 2 серии, превращаем в датафрейм, меняем названия колонок
#     # general_group = group_dpo.concat(group_po)
#     general_group = pd.concat([group_dpo,group_po])
#     general_group = general_group.to_frame().reset_index()
#     general_group.columns = ['Наименование', 'Количество обучающихся по каждому направлению']
#     return general_group





# Загружаем датафреймы
dpo_df = pd.read_excel('data/Тестовая общая таблица.xlsx', sheet_name='ДПО')
po_df = pd.read_excel('data/Тестовая общая таблица.xlsx', sheet_name='ПО')

# Заполняем пустые поля для удобства группировки
dpo_df = dpo_df.fillna('Не заполнено!!!')
po_df = po_df.fillna('Не заполнено!!!')
# Создаем переменную для хранения строки на которой заканчивается предыдущий показатель
border_row = 2
border_column = 2

# Получение общего количества прошедших обучение,количества прошедших по ДПО,по ПО
total_students, total_students_dpo, total_students_po = counting_total_student(dpo_df, po_df)
print(total_students, total_students_dpo, total_students_po)

# Количество обучившихся ДПО и ПО


# Получение количества обучившихся по видам
df_counting_type_and_name_trainning = counting_type_of_training(dpo_df, po_df)
print(df_counting_type_and_name_trainning)

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

# # Добавляем таблицу с по направлениям

sheet['A7'] = 'Вид обучения'
sheet['B7'] = 'Название программы'
sheet['C7'] = 'Количество'

for row in df_counting_type_and_name_trainning.values.tolist():
    sheet.append(row)

sheet.column_dimensions['A'].width = 50
sheet.column_dimensions['B'].width = 30
# Сохраняем файл
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
wb.save(f'Сводный отчет {current_time}.xlsx')

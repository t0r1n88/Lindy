import pandas as pd
import openpyxl

def counting_total_student(lst_df:list):
    """
    Функция для подсчета общего количества студентов обучающихся в цопп
    :param lst_df: список обрабатываемых датафреймов
    :return: общее число обучающихся
    """
    # переменная для подсчета общего количества
    total = 0
    # Перебираем список
    for df in lst_df:
        total +=df.shape[0]
    return total

def counting_type_of_training(lst_df:list):
    """
    Функция для подсчета количества студентов на каждом из видов обучения(повышение квалификации, переподготовка и т.д)
    :param lst_df: список обрабатываемых датафреймов. Нулевой элемент это ДПО, первый это ПО
    :return: датафрейм вида  Вид обучения - количество студентов
    Так как названия колонок в датафреймах отличаются придется обрабатывать по отдельности
    """
    # Обрабатываем ДПО
    dpo = lst_df[0]
    group_dpo = dpo['Дополнительная_профессиональная_программа_повышение_квалификации_профессиональная_переподготовка'].value_counts()

    # Обрабатываем ПО
    po = lst_df[1]
    group_po = po['Программа_профессионального_обучения,_направление_подготовки'].value_counts()
    # Соединяем 2 серии, превращаем в датафрейм, меняем названия колонок
    general_group = group_dpo.append(group_po)
    general_group = general_group.to_frame().reset_index()
    general_group.columns = ['Наименование','Количество обучающихся по каждому направлению']
    return general_group






# Создаем новый excel файл
wb =openpyxl.Workbook()

# Получаем активный лист
sheet = wb.active
sheet.title = 'Сводные данные'

# Сохраняем файл
wb.save('Сводный отчет.xlsx')



# Загружаем датафреймы
dpo_df = pd.read_excel('data/Тестовая общая таблица.xlsx', sheet_name='ДПО')
po_df = pd.read_excel('data/Тестовая общая таблица.xlsx',sheet_name='ПО')

# Заполняем пустые поля для удобства группировки
dpo_df = dpo_df.fillna('Не заполнено!!!')
po_df = po_df.fillna('Не заполнено!!!')
lst_df = [dpo_df,po_df]
# Создаем переменную для хранения строки на которой заканчивается предыдущий показатель
border_row = 2
border_column = 2

# Получение общего количества прошедших обучение
# total_students = counting_total_student(lst_df)
# print(total_students)

# Получение количества обучившихся по видам
df_counting_type_of_trainning = counting_type_of_training(lst_df)
print(df_counting_type_of_trainning)






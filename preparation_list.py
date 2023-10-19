"""
Скрипт для подготовки списка
Очистка некорректных данных, удаление лишних пробелов
"""
import pandas as pd
import openpyxl
import datetime
import re

def create_doc_convert_date(cell):
    """
    Функция для конвертации даты при создании документов
    :param cell:
    :return:
    """
    try:
        string_date = datetime.datetime.strftime(cell, '%d.%m.%Y')
        return string_date
    except ValueError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'
    except TypeError:
        return 'Не удалось конвертировать дату.Проверьте значение ячейки!!!'


def capitalize_fio(value:str)->str:
    """
    Функция для применения capitalize к значениям состоящим из несколько слов разделенных пробелами
    value: значение ячейки
    """
    value = str(value)
    temp_lst = value.split(' ') # создаем список по пробелу
    temp_lst = list(map(str.capitalize,temp_lst))  # обрабатываем
    return ' '.join(temp_lst) #соединяем в строку


def prepare_fio_text_columns(df:pd.DataFrame,lst_columns:list)->pd.DataFrame:
    """
    Функция для очистки текстовых колонок c данными ФИО
    df: датафрейм для обработки
    lst_columns: список колонок которые нужно обработать
    """
    prepared_columns_lst = [] # список для колонок содержащих слова Фамилия,Имя,Отчество, ФИО
    for fio_column in lst_columns:
        for name_column in df.columns:
            if fio_column in name_column.lower():
                prepared_columns_lst.append(name_column)
    if len(prepared_columns_lst) == 0: # проверка на случай не найденных значений
        return df

    df[prepared_columns_lst] = df[prepared_columns_lst].fillna('Не заполнено')
    df[prepared_columns_lst] = df[prepared_columns_lst].astype(str)
    df[prepared_columns_lst] = df[prepared_columns_lst].applymap(lambda x: x.strip() if isinstance(x, str) else x)  # применяем strip, чтобы все данные корректно вставлялись
    df[prepared_columns_lst] = df[prepared_columns_lst].applymap(lambda x:' '.join(x.split())) # убираем лишние пробелы между словами
    df[prepared_columns_lst] = df[prepared_columns_lst].applymap(capitalize_fio)  # делаем заглавными первые буквы слов а остальыне строчными
    return df

def prepare_date_column(df:pd.DataFrame,lst_columns:list)->pd.DataFrame:
    """
    Функция для обработки колонок с датами
    df: датафрейм для обработки
    lst_columns: список колонок которые нужно обработать
    """
    prepared_columns_lst = [] # список для колонок содержащих слово дата
    for date_column in lst_columns:
        for name_column in df.columns:
            if date_column in name_column.lower():
                prepared_columns_lst.append(name_column)
    if len(prepared_columns_lst) == 0: # проверка на случай не найденных значений
        return df

    df[prepared_columns_lst] = df[prepared_columns_lst].applymap(lambda x:pd.to_datetime(x,errors='coerce',dayfirst=True)) # приводим к типу дата
    df[prepared_columns_lst] = df[prepared_columns_lst].applymap(create_doc_convert_date) # приводим к виду ДД.ММ.ГГГГ
    return df

def prepare_snils(df:pd.DataFrame,snils:str)->pd.DataFrame:
    """
    Функция для обработки колонок со снилс
    df: датафрейм для обработки
    snils: название снилс
    """

    prepared_columns_lst = []  # список для колонок содержащих слово дата
    for name_column in df.columns:
        if snils in name_column.lower():
            prepared_columns_lst.append(name_column)

    if len(prepared_columns_lst) == 0: # проверка на случай не найденных значений
        return df

    df[prepared_columns_lst] = df[prepared_columns_lst].applymap(check_snils)
    print(df[prepared_columns_lst].columns)
    print(df[prepared_columns_lst])
    return df





def check_snils(snils):
    """
    Функция для приведения значений снилс в вид ХХХ-ХХХ-ХХХ ХХ
    """
    snils = str(snils)
    result = re.findall(r'\d', snils) # ищем цифры
    if len(result) == 11:
        first_group = ''.join(result[:3])
        second_group = ''.join(result[3:6])
        third_group = ''.join(result[6:9])
        four_group = ''.join(result[9:11])

        out_snils = f'{first_group}-{second_group}-{third_group} {four_group}'
        return out_snils
    else:
        return f'Неправильное значение СНИЛС {snils}'


def prepare_list(file_data:str,path_end_folder:str):
    """
    file_data : путь к файлу который нужно преобразовать
    path_end_folder :  путь к конечной папке
    """
    df = pd.read_excel(file_data,dtype=str) # считываем датафрейм
    df.columns = list(map(str,list(df.columns))) # делаем названия колонок строкововыми
    # обрабатываем колонки с фио
    part_fio_columns = ['фамилия','имя','отчество','фио'] # колонки с типичными названиями
    df = prepare_fio_text_columns(df,part_fio_columns) # очищаем колонки с фио

    # обрабатываем колонки содержащими слово дата
    part_date_columns = ['дата']
    df = prepare_date_column(df,part_date_columns)

    # обрабатываем колонки со снилс
    snils = 'снилс'
    df = prepare_snils(df, snils)



    df.to_excel(f'{path_end_folder}/Датафрейм.xlsx',index=False)




if __name__ == '__main__':
    file_data_main = 'data/example/файл с яндекса.xlsx'
    path_end_main = 'data/example'
    prepare_list(file_data_main,path_end_main)

    print('Lindy Booth')


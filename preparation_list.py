"""
Скрипт для подготовки списка
Очистка некорректных данных, удаление лишних пробелов
"""
import pandas as pd
import openpyxl
import datetime

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


def prepare_fio_text_columns(df:pd.DataFrame,lst_columns:list)->pd.DataFrame:
    """
    Функция для очистки текстовых колонок c данными ФИО
    df: датафрейм для обработки
    lst_columns: список колонок которые нужно обработать
    """
    df[lst_columns] = df[lst_columns].fillna('Не заполнено')
    df[lst_columns] = df[lst_columns].astype(str)
    df[lst_columns] = df[lst_columns].applymap(lambda x: x.strip() if isinstance(x, str) else x)  # применяем strip, чтобы все данные корректно вставлялись
    df[lst_columns] = df[lst_columns].applymap(lambda x:' '.join(x.split())) # убираем лишние пробелы между словами
    df[lst_columns] = df[lst_columns].applymap(lambda x:x.capitalize())  # делаем заглавными первые буквы а остальыне строчными
    return df

def prepare_list(file_data:str,path_end_folder:str):
    """
    file_data : путь к файлу который нужно преобразовать
    path_end_folder :  путь к конечной папке
    """
    df = pd.read_excel(file_data,dtype=str)
    part_fio_columns = ['Фамилия','Имя','Отчество']

    df = prepare_fio_text_columns(df,part_fio_columns) # очищаем колонки с фио
    part_date_columns = ['Дата рождения','Дата выдачи паспорта']
    df[part_date_columns] = df[part_date_columns].applymap(lambda x:pd.to_datetime(x,errors='coerce',dayfirst=True))
    df[part_date_columns] = df[part_date_columns].applymap(create_doc_convert_date) # приводим к виду ДД.ММ.ГГГГ


    print(df[part_date_columns])



    df.to_excel(f'{path_end_folder}/Датафрейм.xlsx',index=False)




if __name__ == '__main__':
    file_data_main = 'data/example/файл с яндекса.xlsx'
    path_end_main = 'data/example'
    prepare_list(file_data_main,path_end_main)

    print('Lindy Booth')


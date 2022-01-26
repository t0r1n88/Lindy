import pandas as pd
import os

"""
Создание файлов Логины и пароли академиии в формате excel из csv"""

path = 'c:/Users/1\YandexDisk/ЦОПП/Цифровая платформа/Облако/Данные пользователей/БРИТ/2021/Студенты/Логины и пароли студентов Академия/'

path_excel = 'c:/Users/1\YandexDisk/ЦОПП/Цифровая платформа/Облако/Данные пользователей/БРИТ/2021/Студенты/Excel Логины и пароли студентов Академия/'
for file in os.listdir(path):
    name_file = file.split('.')[0]
    df = pd.read_csv(f'{path}{file}',encoding='cp1251',delimiter=';')
    df.drop(['Страна','Тип'],inplace=True,axis=1)
    df.rename(columns={'email':'Логин'})
    df.to_excel(f'{path_excel}{name_file}.xlsx',index=False)
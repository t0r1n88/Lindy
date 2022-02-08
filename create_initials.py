import pandas as pd

def create_initials(fio):
    """
    Функция для создания инициалов для использования в договорах
    формат фио -Будаев Олег Тимурович
    """
    #Создаем 3 переменные

    initials_firstname = ''
    initials_middlename = ''
    initials_lastname = ''

    # Сплитим по пробелу
    lst_fio = fio.split()
    # Если ФИО стандартное
    if len(lst_fio) == 3:

        lastname = lst_fio[0]
        firstname = lst_fio[1]
        middlename = lst_fio[2]
        # Создаем инициалы
        initials_firstname = firstname[0].upper()
        initials_middlename = middlename[0].upper()
        initials_lastname = lastname
        # Возвращаем полученную строку
        print(f'{initials_firstname}.{initials_middlename} {initials_lastname}')
        return f'{initials_firstname}.{initials_middlename} {initials_lastname}'
    elif len(lst_fio) == 2:
        lastname = lst_fio[0]
        firstname = lst_fio[1]


        initials_firstname = firstname[0].upper()
        initials_lastname = lastname
        print(f'{initials_firstname}. {initials_lastname}')
        return f'{initials_firstname}. {initials_lastname}'
    elif len(lst_fio) == 1:
        lastname = lst_fio[0]
        initials_lastname = lastname
        print(f'{initials_lastname}')
        return f'{initials_lastname}'
    else:
        print('Проверьте правильность написания ФИО в столбце ФИО_именительный\n')








df = pd.read_excel('data/Отдельные таблицы/Тестовая таблица.xlsx')
df['Инициалы'] = df['ФИО_именительный'].apply(create_initials)



wb = openpyxl.load_workbook(name_file_data_obraz_program_po)

name_sheet_up = wb.sheetnames[0]  # получаем название листа с учебным планом
name_sheet_data = wb.sheetnames[1]  # получаем название листа с данными программы

"""
1) Ищем на какой строке находится ИТОГО
2) Мы знаем что там должно быть 7 колонок
"""
target_value = 'ИТОГО'

# Поиск значения в выбранном столбце
column_number = 1  # Номер столбца, в котором ищем значение (например, столбец A)
target_row = None  # Номер строки с искомым значением

for row in wb[name_sheet_up].iter_rows(min_row=1, min_col=column_number, max_col=column_number):
    cell_value = row[0].value
    if cell_value == target_value:
        target_row = row[0].row
        break

if not target_row:
    # если не находим слово ИТОГО то выдаем исключение
    raise NotTotal

# если значение найдено то считываем нужное количество строк и  7 колонок
df_up = pd.read_excel(name_file_data_obraz_program_po, sheet_name=name_sheet_up, nrows=target_row,
                      usecols='A:G', dtype=str)

df_up.iloc[:, 1:6] = df_up.iloc[:, 1:6].applymap(convert_to_int)  # 1) Приводим к инту колонки 2-6

# Заполняем возможные пустые строки
df_up['Наименование_раздела'] = df_up['Наименование_раздела'].fillna('Не заполнено название раздела')
# Очищаем от возможнных пробелов
df_up['Наименование_раздела'] = df_up['Наименование_раздела'].apply(lambda x: x.strip())

# Создаем датафрейм учебной программы без учета строки ИТОГО для таблиц краткой аннотации
short_df_up = df_up[df_up['Наименование_раздела'] != 'ИТОГО']
short_df_up = short_df_up[short_df_up['Наименование_раздела'] != 'Итоговая аттестация']

# получаем единичные значения из листа с данными
single_row_df = pd.read_excel(name_file_data_obraz_program_po, sheet_name=name_sheet_data, nrows=1,
                              usecols='A:K')
single_row_df.iloc[:, 8] = single_row_df.iloc[:, 8].apply(convert_date)  # обрабатываем колонку с датой

# Очищаем от лишнего поля которые заполняет пользователь
# Заполняем возможные пустые строки
single_row_df['Наименование_профессии'] = single_row_df['Наименование_профессии'].fillna('Не заполнено !!!')
# Очищаем от возможнных пробелов
single_row_df['Наименование_профессии'] = single_row_df['Наименование_профессии'].apply(lambda x: x.strip())

single_row_df['Профессиональный_стандарт'] = single_row_df['Профессиональный_стандарт'].fillna(
    'Не заполнено !!!')
# Очищаем от возможнных пробелов
single_row_df['Профессиональный_стандарт'] = single_row_df['Профессиональный_стандарт'].apply(
    lambda x: x.strip())

# получаем датафрейм с технологиями обучения
tech_df = pd.read_excel(name_file_data_obraz_program_po, sheet_name=name_sheet_data, usecols='L:O')

tech_df.dropna(thresh=2, inplace=True)  # очищаем от строк в которых не заполнены 2 колонки

tech_df['Разработчики_программы'] = tech_df['Разработчики_программы'].fillna('Не заполнено')
# Очищаем от возможнных пробелов
tech_df['Характеристика_технологии_обучения'] = tech_df['Характеристика_технологии_обучения'].apply(
    lambda x: x.strip())
tech_df['Технологии_обучения'] = tech_df['Технологии_обучения'].apply(lambda x: x.strip())
tech_df['Разработчики_программы'] = tech_df['Разработчики_программы'].apply(lambda x: x.strip())

tech_df['Уровни_квалификации'] = tech_df['Уровни_квалификации'].fillna(0)
tech_df['Уровни_квалификации'] = tech_df['Уровни_квалификации'].astype(int)

# создаем переменную для уровней квалификации
educ_lst = tech_df['Технологии_обучения'].tolist()

levels_qual = tech_df['Уровни_квалификации'].to_list()
levels_qual = list(filter(lambda x: x != 0, levels_qual))
levels_qual = list(map(str, levels_qual))

# Конвертируем датафрейм с описанием программы в список словарей
data_program = single_row_df.to_dict('records')

context = data_program[0]
# текстовые составные переменные
context['Уровни_квалификации'] = ','.join(levels_qual)
context['Технологии_обучения'] = ';\n'.join(educ_lst)

# Добавляем датафреймы
context['lst_tech'] = tech_df.to_dict('records')  # образовательные технологии
context['up_lst'] = df_up.to_dict('records')  # учебный план
context['short_up_lst'] = short_df_up.to_dict('records')  # учебный план

lst_dev = [value for value in tech_df['Разработчики_программы'].tolist() if value != 'Не заполнено']
context['lst_dev'] = lst_dev

doc = DocxTemplate(name_file_template_educ_program_po)
# Создаем документ
doc.render(context)
# сохраняем документ
# название программы
name_prof = single_row_df['Наименование_профессии'].tolist()[0]
razr = single_row_df['Разряд'].tolist()[0]
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)
doc.save(
    f'{path_to_end_folder_obraz_program_po}/Программа ПО {name_prof} {razr} разряда {current_time}.docx')
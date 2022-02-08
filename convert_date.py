import pandas as pd
import datetime
from docxtpl import DocxTemplate

def convert_date(cell):
    """
    Функция для конвертации даты в формате 1957-05-10 в формат 10.05.1957(строковый)
    """
    string_date = datetime.datetime.strftime(cell,'%d.%m.%Y')
    return string_date

# name_file_data_doc = pd.read_excel('data/Отдельные таблицы/ДПО_Профессиональное_самоопределение_январь.xlsx',dtype={'Дата_рождения_получателя':str})
df = pd.read_excel('data/Отдельные таблицы/ДПО_Профессиональное_самоопределение_январь.xlsx')
df['Дата_рождения_получателя'] = df['Дата_рождения_получателя'].apply(convert_date)
df['Дата_выдачи_паспорта'] = df['Дата_выдачи_паспорта'].apply(convert_date)

print(df['Дата_рождения_получателя'])

data = df.to_dict('records')

for row in data:
    context = row

    doc = DocxTemplate('data/Шаблон_Заявление и согласие на обработку 08.02.docx')
    # Создаем документ
    doc.render(context)
# сохраняем документ
    doc.save(f'Проба времени.docx')
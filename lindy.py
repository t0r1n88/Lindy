import pandas as pd
from docxtpl import DocxTemplate


df = pd.read_excel('data/Тестовая база ДПО.xlsx')
print(df.head())


data = df.to_dict('records')
print(data)

for row in data:
    doc = DocxTemplate('data/Короткий тестовый шаблон ДПО.docx')
    context = row
    # Превращаем строку в список кортежей, где первый элемент кортежа это ключ а второй данные
    id_row = list(row.items())


    print(id_row[0][1])
    doc.render(context)
    doc.save(f'{id_row[0][1]}.docx')



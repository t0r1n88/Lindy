import pandas as pd
from docxtpl import DocxTemplate





df = pd.read_excel('data/Форма базы данных.xlsx')
print(df.head())


data = df.to_dict('records')

lst_students = []

for row in data:
    doc = DocxTemplate('data/Форма Приказ о зачислении_ДПО.docx')
    context = row
    fio = row['ФИО_именительный']
    lst_students.append(fio)

    # Превращаем строку в список кортежей, где первый элемент кортежа это ключ а второй данные
    # id_row = list(row.items())

    doc.render(context)
    doc.save(f'{fio}.docx')

doc = DocxTemplate('data/Форма Приказ о зачислении_ДПО.docx')



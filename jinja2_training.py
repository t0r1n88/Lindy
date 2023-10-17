import pandas as pd
from docxtpl import DocxTemplate
import time

template = 'jinja2_train.docx'
data_table = 'first.xlsx'

df = pd.read_excel(data_table)

data = df.to_dict('records')

doc = DocxTemplate(template)
# context = data[0]
context = dict()
context['FIO'] = df['ФИО']
context['Series_pas'] = df['Серия_паспорта']
# # заполняем пустые значения в столбце
# avg_df = df['Средний_балл']
# avg_df.fillna(0,inplace=True)
#
# df['Средний_балл'] = df['Средний_балл'].astype('int')
context['AVG'] = df['Средний_балл']

doc.render(context)
t = time.localtime()
current_time = time.strftime('%H_%M_%S', t)

doc.save(
    f'{current_time}.docx')


import pandas as pd
import openpyxl

df = pd.read_excel('data/колонки ПО.xlsx')

df['Название'] = df['Название'].apply(lambda x: '{{' + f'{x}' + '}}')
df.to_excel('Для тестового шаблона ПО.xlsx',index=False)




df = pd.read_excel('data/колонки ДПО.xlsx')

df['Название'] = df['Название'].apply(lambda x: '{{' + f'{x}' + '}}')
df.to_excel('Для тестового шаблона ДПО.xlsx',index=False)
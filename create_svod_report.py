import pandas as pd
import openpyxl

# Загружаем лист ДПО
dpo_df = pd.read_excel('data/Форма базы данных.xlsx', sheet_name='ДПО')
print(dpo_df)
po_df = pd.read_excel('data/Форма базы данных.xlsx',sheet_name='ПО')
print(po_df)

# Получение общего количества прошедших обучение

# Получение количества обучившихся

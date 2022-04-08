"""
Скрипт для поиска совпадающих значений из  2 столбцов
"""
import pandas as pd
import openpyxl

big_df = pd.read_excel('big.xlsx')
small_df = pd.read_excel('small.xlsx')

# itog_df = big_df.join(small_df,on='ФИО',how='inner')
itog_df = pd.merge(small_df,big_df,how='inner')
itog_df.to_excel('Общее.xlsx',index=False)

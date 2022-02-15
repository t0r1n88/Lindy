import pandas as pd




name_file_data_report = 'Общая таблица слушателей ЦОПП.xlsx'

df_dpo = pd.read_excel(name_file_data_report,sheet_name='ДПО')

group_quantity_program = df_dpo.groupby(['Наименование_дополнительной_профессиональной_программы'])

print(group_quantity_program)

# dpo_age_distribution = pd.pivot_table(df_dpo, index=['Дополнительная_профессиональная_программа_повышение_квалификации_профессиональная_переподготовка'],
#                                       values=['Дополнительная_профессиональная_программа_повышение_квалификации_профессиональная_переподготовка'],
#                                       aggfunc={'Дополнительная_профессиональная_программа_повышение_квалификации_профессиональная_переподготовка':'count'})
# dpo_age_distribution.to_excel('Первый шаг.xlsx')
# dpo_age_distribution = dpo_age_distribution.reset_index()
# dpo_age_distribution.to_excel('После ресета.xlsx')

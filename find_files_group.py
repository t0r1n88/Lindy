import os
import re
# path_to_end_folder_doc = filedialog.askdirectory()

# Комплируем регулярное выражение
pattern = re.compile('^[А-ЯЁ]+_.+_(?:январь|февраль|март|апрель|май|июнь|июль|август|сентябрь|октябрь|ноябрь|декабрь)\.xlsx$')
path_to_files_groups = 'z:/!!!БАЗА ДАННЫХ/2022/'




for dirpath,dirnames,filenames in os.walk(path_to_files_groups):
    # перебрать файлы
    for filename in filenames:
        if re.search(pattern,filename):
            print("Файл:", os.path.join(dirpath, filename))


from tkinter import *
from tkinter import ttk


class Notebook:



    def __init__(self,title):
        self.root = Tk()
        self.root.title(title)
        self.notebook = ttk.Notebook(self.root)



    def add_tab(self,title,text):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame,text=title)
        label = ttk.Label(frame,text=text)
        label.grid(column=1,row=1)
        self.notebook.pack()



    def run(self):
        self.root.mainloop()



nb = Notebook('Example')
nb.add_tab('Frame One','This is on Frame One')
nb.add_tab('Frame Two','This is on Frame Two')
nb.run()

# # Считываем csv файл, не забывая что екселевский csv разделен на самомо деле не запятыми а точкой с запятой
#         reader = csv.DictReader(open(name_file_data_doc), delimiter=';')
#         # Конвертируем объект reader в список словарей
#         data = list(reader)
#         # Создаем в цикле документы
#         for row in data:
#             doc = DocxTemplate(name_file_template_doc)
#             context = row
#             # Превращаем строку в список кортежей, где первый элемент кортежа это ключ а второй данные
#             id_row = list(row.items())
#             try:
#                 doc.render(context)
#                 print(context)
#                 print(id_row[0][1])
#                 doc.save(f'{path_to_end_folder_doc}/{id_row[0][1]}.docx')
#             except:
#                 messagebox.showerror('Dodger','Проверьте содержимое шаблона\nНе допускаются пробелы в словах внутри фигурных скобок')
#                 continue
#         messagebox.showinfo('Dodger', 'Создание документов успешно завершено!')
#     except NameError as e:
#         messagebox.showinfo('Dodger', f'Выберите шаблон,файл с данными и папку куда будут генерироваться файлы')
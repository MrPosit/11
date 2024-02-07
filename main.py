'''а) Прочитайте из трёх excel файлов (заранее создайте их, внутри 1111, 2222, 3333).
б) Отсортируйте полученную матрицу в порядке убывания.
в) Запишите это в один файл, изменив шрифт и обернув в границы.'''

# import pandas as pd

# excel_file1 = pd.ExcelFile("1111.xlsx")
# df1 = pd.read_excel(excel_file1, engine='xlrd')
# excel_file2 = pd.ExcelFile("2222.xlsx")
# df2 = pd.read_excel(excel_file2, engine='xlrd')
# excel_file3 = pd.ExcelFile("3333.xlsx")
# df3 = pd.read_excel(excel_file3, engine='xlrd')

# combined_df = pd.concat([df1, df2, df3], ignore_index=True)
# sorted_df = combined_df.sort_values(by=combined_df.columns[0], ascending=False)

# writer = pd.ExcelWriter("combined_sorted.xlsx", engine='xlsxwriter')
# sorted_df.to_excel(writer, index=False)
# workbook = writer.book
# worksheet = writer.sheets['Sheet1']

# formatting = workbook.add_format({'border': 1, 'font_name': 'Arial'})
# for col_num, value in enumerate(sorted_df.columns.values):
#     worksheet.write(0, col_num, value, formatting)

# for row_num in range(1, len(sorted_df) + 1):
#     for col_num, value in enumerate(sorted_df.iloc[row_num - 1]):
#         worksheet.write(row_num, col_num, value, formatting)

# writer.save()



'''а) Создайте json файл в операционной системе, заполните его данными с сайта
https://jsonplaceholder.typicode.com/todos/
б) Прочитайте этот файл в массив python-словарей.
в) Запишите каждый из этих словарей в отдельный json-файл.'''
# import json
# import requests

# url = 'https://jsonplaceholder.typicode.com/todos/'
# todos = requests.get(url).json()

# for idx, item in enumerate(todos, start=1):
#     with open(f'todo_{idx}.json', 'w') as file:
#         json.dump(item, file, indent=4)


'''а) Создайте word файл в операционной системе, заполните его текстом «Hello Python».
б) Прочитайте этот файл, только жирный текст в python-строку и выведите на экран.
в) Создайте абзац с текстом и запишите это в новый word-файл, изменит шрифт и размер
шрифта.
'''

# from docx import Document
# from docx.shared import Pt

# doc = Document()
# doc.add_paragraph("Hello Python")
# doc.save("hello_python.docx")

# bold_text = ""
# doc = Document("hello_python.docx")
# for paragraph in doc.paragraphs:
#     for run in paragraph.runs:
#         if run.bold:
#             bold_text += run.text

# doc = Document()
# paragraph = doc.add_paragraph("Это абзац текста.")
# font = paragraph.runs[0].font
# font.name = 'Times New Roman'
# font.size = Pt(12)
# doc.save("formatted_text.docx")

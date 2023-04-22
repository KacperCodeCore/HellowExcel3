import excel_operations as excel_o

from _ast import Interactive

import openpyxl as xl
import win32com.client
import os

# Convert csv to Excel and save
import pandas as pd
df = pd.read_csv('Book1.csv', encoding='Windows-1250', delimiter=';')
df.to_excel('Book1.xlsx', index=False, header=True)



# Load the Excel workbook
work_book = xl.load_workbook('Book1.xlsx')
sheet = work_book.active

# excel_o.add_rows(work_book, sheet)
# work_book.save('Book1Done.xlsx')
# excel_o.save_excel_to_pdf('Book1Done.xlsx')

file_path = os.path.join(os.getcwd(), 'Book1Done.xlsx')
print(file_path)

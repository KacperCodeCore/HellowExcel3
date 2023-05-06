import excel_operations as excel

from _ast import Interactive

import openpyxl as xl
import win32com.client
import os
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Convert csv to Excel and save
import pandas as pd
df = pd.read_csv('Book1.csv', encoding='Windows-1250', delimiter=';')
df.to_excel('Book1.xlsx', index=False, header=True)



# Load the Excel workbook
work_book = xl.load_workbook('Book1.xlsx')
sheet = work_book.active

excel.add_rows(work_book, sheet)
work_book.save('Book1Done.xlsx')
#excel.save_excel_to_pdf('Book1Done.xlsx')

# usówanie kolumn i formatowanie jako tabela
# work_book = xl.load_workbook('Book1Done.xlsx')
# sheet = work_book.active
# excel.remove_columns_except(work_book, sheet)
excel.format_excel_file('Book1Done.xlsx')
excel.save_excel_to_pdf('Book1Book1Done.xlsx')# !!! trzeba dodać zapisywanie z dowolnj ścieżki




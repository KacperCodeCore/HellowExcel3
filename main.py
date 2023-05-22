import excel_operations as excel
import openpyxl as xl

# Convert csv to Excel and save
import pandas as pd
df = pd.read_csv('Book1.csv', encoding='Windows-1250', delimiter=';')
df.to_excel('Book1.xlsx', index=False, header=True)

# Load the Excel workbook
work_book = xl.load_workbook('Book1.xlsx')
sheet = work_book.active

excel.add_rows(work_book)
# work_book.save('Book1Done.xlsx')
excel.sub_collumns(work_book)
excel.save_excel_to_pdf('Book1.xlsx')# !!! trzeba dodać zapisywanie z dowolnj ścieżki




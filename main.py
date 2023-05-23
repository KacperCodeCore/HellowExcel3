import os
import excel_operations as excel
import openpyxl as xl
import pandas as pd


# Find all .cvs files
current_directory = os.getcwd()
file_list = os.listdir(current_directory)

for file_name in file_list:
    if file_name.endswith('.csv'):

        # Convert csv to Excel and save
        df = pd.read_csv(file_name, encoding='Windows-1250', delimiter=';')
        df.to_excel('Book1__________________Operations.xlsx', index=False, header=True)

        # Load the Excel workbook
        work_book = xl.load_workbook('Book1__________________Operations.xlsx')
        sheet = work_book.active

        # Excel operations
        excel.add_rows(work_book)
        excel.sub_collumns(work_book)

        # remove extension from filename
        file_name_without_ext = file_name.split('.')[0]

        excel.save_excel_to_pdf('Book1__________________Operations.xlsx', file_name_without_ext)
        excel.delete_file('Book1__________________Operations.xlsx')





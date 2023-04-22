
import win32com.client
import os

def add_rows(work_book, sheet):
    # Add new rows and set correct values
    rowIndex = 2
    rowMax = sheet.max_row + 1
    while rowIndex < rowMax:
        cell = sheet.cell(rowIndex, 6).value
        if cell == 28:
            # Add new row
            sheet.insert_rows(rowIndex + 1)
            sheet.cell(rowIndex, 6).value = 10
            # Operations on operation text.1 for current row
            operation_text = str(sheet.cell(rowIndex, 16).value)
            operation_text = operation_text[:3] + "10" + operation_text[5:]
            sheet.cell(rowIndex, 16).value = operation_text
            # Copy values from row above
            for column in range(1, sheet.max_column + 1):
                sheet.cell(row=rowIndex + 1, column=column).value = sheet.cell(row=rowIndex, column=column).value
            sheet.cell(rowIndex + 1, 6).value = 18
            # Operations on operation text.1 for new row
            operation_text = str(sheet.cell(rowIndex + 1, 16).value)
            operation_text = operation_text[:3] + "18" + operation_text[5:]
            sheet.cell(rowIndex + 1, 16).value = operation_text
            #
            rowIndex += 1
            rowMax += 1
        if cell == 36:
            # Add new row
            sheet.insert_rows(rowIndex + 1)
            sheet.cell(rowIndex, 6).value = 18
            # Operations on operation text.1 for current row
            operation_text = str(sheet.cell(rowIndex, 16).value)
            operation_text = operation_text[:3] + "18" + operation_text[5:]
            sheet.cell(rowIndex, 16).value = operation_text
            # Copy values from row above
            for column in range(1, sheet.max_column + 1):
                sheet.cell(row=rowIndex + 1, column=column).value = sheet.cell(row=rowIndex, column=column).value
            #
            rowIndex += 1
            rowMax += 1
        rowIndex += 1  # Move to the newly inserted row




def save_excel_to_pdf(file_name):
    file_path = os.path.join(os.getcwd(), file_name)

    # Convert Excel file to PDF using win32com
    excel = win32com.client.Dispatch('Excel.Application')
    workbook = excel.Workbooks.Open(r'D:\GitHub\Python\HellowExcel3\Book1Done.xlsx')
    workbook.ExportAsFixedFormat(0, r'D:\GitHub\Python\HellowExcel3\Book1Done.pdf', 1, 0, 0)

    # Close the Excel file
    workbook.Close()
    excel.Quit()
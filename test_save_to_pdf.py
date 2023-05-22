# import openpyxl as xl
#
# # Load the Excel workbook
# wb = xl.load_workbook('Book1Done.xlsx')
# sheet = wb.active
#
# # Add table and format it
# table = sheet["J1"].expand("table")
# if isinstance(table, tuple):
#     table = table[1:]
#
# table_name = "TABELA_2"
# sheet.tables.add(table, table_name)
# sheet.tables[table_name].table_style = "TableStyleLight18"
#
# # Auto fit all columns
# for col in sheet.columns:
#     max_length = 0
#     column = col[0].column_letter  # Get the column name
#     for cell in col:
#         try:
#             if len(str(cell.value)) > max_length:
#                 max_length = len(str(cell.value))
#         except:
#             pass
#     adjusted_width = (max_length + 2) * 1.2  # Adjust the width for padding
#     sheet.column_dimensions[column].width = adjusted_width
#
# # Insert row and copy values
# sheet.insert_rows(1)
# sheet["R3"].copy(sheet["A1"])
# sheet["U3"].copy(sheet["E1"])
# sheet["V3"].copy(sheet["Q1"])
#
# # Delete unnecessary columns
# sheet.delete_cols(4, 1)  # Delete column D
# sheet.delete_cols(6, 8)  # Delete columns F to O
# sheet.delete_cols(8, 5)  # Delete columns G to L
#
# # Merge cells and format headers
# sheet.merge_cells("A1:C1")
# sheet["A1"].alignment = xl.styles.Alignment(horizontal="center", vertical="bottom")
# sheet["D1"].alignment = xl.styles.Alignment(horizontal="center", vertical="bottom")
# sheet["A1"].font = xl.styles.Font(bold=True)
# sheet["D1"].font = xl.styles.Font(bold=True)
#
# # Save the Excel file
#
# wb.save('Book1Done2Done.xlsx')
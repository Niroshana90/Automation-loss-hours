import openpyxl
wb = openpyxl.load_workbook('New Microsoft Excel Worksheet.xlsx')
print(wb.get_sheet_names())
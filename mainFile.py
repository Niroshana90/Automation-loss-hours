# adding file save
import openpyxl
import os

source_file = 'Autonomation update Central team.xlsx'
result_file = 'result.xlsx'

if not os.path.isfile(source_file):
    print('source file dose not exists')
    exit()
else:
    print('source file exists')

wb = openpyxl.load_workbook(source_file)
print(wb.get_sheet_names().__len__())

if os.path.isfile(result_file):
    os.remove(result_file)
    print("removing existing files")

print("creating new result file")
result_wb = openpyxl.Workbook()
result_wb.save(result_file)
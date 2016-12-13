# adding file save
import openpyxl
import os

source_file = 'BW22 Nov.xlsx'
result_file = 'result.xlsx'

if not os.path.isfile(source_file):
    print('source file dose not exists')
    exit()
else:
    print('source file exists')

wb = openpyxl.load_workbook(source_file)
ws = wb.get_active_sheet()

print(ws['A'][3].value+'\t\t'+ws['B'][3].value+'\t\t'+ws['F'][2].value+'\t\t'+ws['G'][2].value+'\t\t'+ws['H'][2].value+'\t\t'+ws['I'][2].value+'\t\t'+ws['J'][2].value+'\t\t'+ws['K'][2].value+'\t\t')


for i in range(1, ws.max_row - ws.min_row):
    if ws['B'][i].value == 'Result':
        print(str(ws['A'][i].value) + '\t\t' + ws['B'][i].value + '\t\t' + str(ws['F'][i].value) + '\t\t' + str(ws['G'][i].value) + '\t\t' + str(ws['H'][i].value) + '\t\t' + str(ws['I'][i].value) + '\t\t' + str(ws['J'][i].value) + '\t\t' + str(ws['K'][i].value) )

if os.path.isfile(result_file):
    os.remove(result_file)
    print("removing existing files")

print("creating new result file")
result_wb = openpyxl.Workbook()
result_wb.save(result_file)
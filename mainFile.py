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
if os.path.isfile(result_file):
    os.remove(result_file)
    print("removing existing files")

print("creating new result file")
result_wb = openpyxl.Workbook()
written_row_count = int(0)
result_wb_ws = result_wb.get_active_sheet()
result_wb_ws['A'][0].value = ws['A'][3].value
result_wb_ws['B'][0].value = ws['B'][3].value
result_wb_ws['C'][0].value = ws['F'][2].value
result_wb_ws['D'][0].value = ws['G'][2].value
result_wb_ws['E'][0].value = ws['H'][2].value
result_wb_ws['F'][0].value = ws['I'][2].value
result_wb_ws['G'][0].value = ws['J'][2].value
result_wb.save(result_file)
written_row_count = 2

for i in range(1, ws.max_row - ws.min_row):
    if ws['B'][i].value == 'Result':
        result_wb_ws.cell(row=written_row_count, column=1).value = ws['A'][i].value
        result_wb_ws.cell(row=written_row_count, column=2).value = ws['B'][i].value
        result_wb_ws.cell(row=written_row_count, column=3).value = ws['F'][i].value
        result_wb_ws.cell(row=written_row_count, column=4).value = ws['G'][i].value
        result_wb_ws.cell(row=written_row_count, column=5).value = ws['H'][i].value
        result_wb_ws.cell(row=written_row_count, column=6).value = ws['I'][i].value
        result_wb_ws.cell(row=written_row_count, column=7).value = ws['J'][i].value
        print('.', sep=' ', end='', flush=True)
        written_row_count += 1
        #print(str(ws['A'][i].value) + '\t\t' + ws['B'][i].value + '\t\t' + str(ws['F'][i].value) + '\t\t' + str(ws['G'][i].value) + '\t\t' + str(ws['H'][i].value) + '\t\t' + str(ws['I'][i].value) + '\t\t' + str(ws['J'][i].value) + '\t\t' + str(ws['K'][i].value) )


result_wb.save(result_file)
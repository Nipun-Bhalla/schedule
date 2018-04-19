import openpyxl
import time
import json

localtime = time.asctime(time.localtime(time.time()) )
print (localtime)
day = localtime.split(' ', 1)[0]

workbook = openpyxl.load_workbook('D:/Smart Mirror/SmartMirror/Time Tables/Schedule.xlsx')
worksheet = workbook[day]

init_val = 0
with open('D:/Smart Mirror/SmartMirror/Time Tables/availability.json', mode='w') as outfile:
            json.dump(init_val, outfile)

for row in worksheet.iter_rows(min_row=2, min_col=1, max_col=1):
    for cell in row:
        print(cell.value)
        data = cell.value
        with open('D:/Smart Mirror/SmartMirror/Time Tables/availability.json', mode='a') as outfile:
            json.dump(data, outfile)



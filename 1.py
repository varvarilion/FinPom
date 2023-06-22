from openpyxl import *

wb = load_workbook('FinPom2.xlsx')

sheet2 = wb.worksheets[0]
current_row = sheet2.max_row

sheet2.cell(row=current_row + 1, column=8).value = int(input())  # доход наличные

sheet2.cell(row=current_row + 1, column=9).value = int(input())  # доход безнал

sheet2.cell(row=current_row + 1, column=5).value = int(input())  # расход наличные

sheet2.cell(row=current_row + 1, column=6).value = int(input())  # расход безнал

wb.save('FinPom2.xlsx')

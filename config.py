from openpyxl import *

wb = load_workbook('balances.xlsx')
sheet1 = wb.active
sheet2 = wb.worksheets[1]
sheet3 = wb.worksheets[2]


def excel():
    sheet1.column_dimensions['A'].width = 20
    sheet2.column_dimensions['A'].width = 20
    sheet2.column_dimensions['B'].width = 20
    sheet3.column_dimensions['A'].width = 20
    sheet3.column_dimensions['B'].width = 20


def wordfinder(ABC):
    for i in range(1, sheet1.max_row + 1):
        for j in range(1, sheet1.max_column + 1):
            if ABC == sheet1.cell(i, j).value:
                a = i
                return a

current_row = sheet1.max_row
if wordfinder(str(input("ID"))) == None:
    sheet1.cell(row=current_row + 1, column=1).value = input("ID")
    sheet2.cell(row=current_row + 1, column=1).value = input("Категория:")
    sheet2.cell(row=current_row + 1, column=2).value = int(input("Доход:"))
    sheet3.cell(row=current_row + 1, column=1).value = input("Категория:")
    sheet3.cell(row=current_row + 1, column=2).value = int(input("Расход:"))
else:
    a = wordfinder(str(input("ID")))
    sheet2.cell(row=a, column=1).value = input("Категория:")
    sheet2.cell(row=a, column=2).value = input("Доход:")
    sheet3.cell(row=a, column=1).value = input("Категория:")
    sheet3.cell(row=a, column=2).value = input("Расход:")
wb.save('balances.xlsx')
excel()

wb.save('balances.xlsx')

from openpyxl import *
wb = load_workbook('balances.xlsx')
sheet = wb.active
def excel():
    sheet.column_dimensions['A'].width = 30
    sheet.column_dimensions['B'].width = 10
current_row = sheet.max_row
current_column = sheet.max_column

sheet.cell(row=current_row + 1, column=1).value = input()
sheet.cell(row=current_row + 1, column=2).value = input()
wb.save('balances.xlsx')
if __name__ == "__main__":
    excel()
wb.save('balances.xlsx')

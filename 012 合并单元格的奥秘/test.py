import openpyxl

def get_salary():
    wb = openpyxl.load_workbook('salary.xlsx')
    ws = wb.active
    max_row = ws.max_row()

get_salary()

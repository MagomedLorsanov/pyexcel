from openpyxl import Workbook, load_workbook

wb = load_workbook('./excels/РЕЙТИНГ с изм.2.xlsx')

ws = wb.active

print(ws)
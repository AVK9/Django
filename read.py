import openpyxl as op

filename = 'test.xlsx'

wb = op.load_workbook(filename, data_only=True)
sheet = wb.active

max_row = sheet.max_row

print(max_row)

print(sheet.cell(row=11, column=1).value)
import openpyxl
from openpyxl import load_workbook
wb=load_workbook("automation-book-example.xlsx")

#get the sheet
ws=wb.get_sheet_by_name('Sheet1')

#accessing a single access
a=ws['A1']

#for accessing   a.row it will give 1 2 3 row number
# a.column  it give A B C    a.value give the value in the cell

#a.coordinate print A1


print(ws.max_row)
print(ws.max_column)

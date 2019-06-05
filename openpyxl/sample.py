from openpyxl import load_workbook
wb=load_workbook("sample.xlsx")

#in openpyxl you can actually create a excel document itself
#and save it in xlsx format

#read openpyxl library documentation

#printing the sheet names this is in list format
print(wb.sheetnames)

#getting current active sheet
ws=wb.active

#get by sheet name Sheet1
ws=wb['Sheet1']

#another one
ws=wb.get_sheet_by_name('Sheet1')

#accessing one cell
print(ws['A1'].value)

#maximum rows
print(ws.max_row)

print("\n")
#printing particular column values A B C..
#same for priting row speifiy 1 2 3 ....
for i in ws['A']:
    print(i.value)

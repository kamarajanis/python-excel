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

#for printing rows values and column values
#ws.rows and ws.columns

##print("\n\n\n")
##print("*********************************\n**************************")
###here ws.rows print the tuple row by row
##for i in ws.rows:            #same use ws,colums to print column and next column
##    #here to get the row value
##    for j in i:
##        #it will print the value row by row
##        print(j.value)



##
###to print the values
###ws.values it will print all vallues row by row
##print("*********************************\n**************************")
##for i in ws.values:
##    print(i)#that row is also a tuple
##
##print("*********************************\n**************************")
##print("\n\n\n")
##
##for i in ws.values:
##    for j in i:
##        print(j)

for i in ws.rows:
    print(i)




        
    

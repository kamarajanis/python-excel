from openpyxl import load_workbook
wb=load_workbook("time-write.xlsx")
import datetime
ws=wb.active
for i in range(1,11):
    t=datetime.datetime.now()
    ws.cell(row=i,column=1,value=t)

wb.save('time-write.xlsx')

#you can see the time that are stored in excel sheet 

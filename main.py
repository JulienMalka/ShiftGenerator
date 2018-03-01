import openpyxl

book = openpyxl.load_workbook('shifts.xlsx')
sheet = book.active

NB_ROW = 11
NB_COLUMN = 12

searched = str(112)

for i in range(1, NB_COLUMN):
    for j in range(1, NB_ROW):
        cell = str(sheet.cell(row=j, column=i).value)

        if searched in cell:
            print("Assignment :"+str(sheet.cell(row=j,column=2).value)+"  "+str(sheet.cell(row=1, column=i).value))

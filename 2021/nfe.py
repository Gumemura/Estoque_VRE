import xlrd

f = open("items.xml", "w")

# Give the location of the file
loc = ("Estoque VRE 2021.xlsx")
 
# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_name("Saida")

num_rows= sheet.nrows

products = []
i = 0
while i < 10:
    if sheet.cell_value(1, i) != "":
        lastCol = i
    i = i + 1

i = 0

for i in range(2, num_rows):
    if sheet.cell_value(i, 1) == 0:
        break;
    else:
        if sheet.cell_value(i, lastCol) != "":
            print(sheet.cell_value(i, 0) + "\t\t\t" + str(sheet.cell_value(i, lastCol)))

'''for elem in products:
    print("Name: " + elem[0] + "\t\t| " + "Price: " + str(elem[1]))
    f.write(elem[0] + " " + str(elem[1]) + "\n")'''

f.close()
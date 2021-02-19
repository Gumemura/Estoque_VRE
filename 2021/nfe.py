import xlrd

FIRST_COLUNM = 2
BOARD_CODE = 1
BOARD_NAME = 0

f = open("items.xml", "w")

# Give the location of the file
loc = ("Estoque VRE 2021.xlsx")
 
# To open Workbook
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_name("Saida")

num_rows= sheet.nrows

i = FIRST_COLUNM
lastCol = 0
while i < sheet.ncols:
    if sheet.cell_value(1, i) != "":
        lastCol = i
    else:
        break
    i = i + 1

i = 0
products = []
for i in range(2, num_rows):
    if sheet.cell_value(i, BOARD_CODE) == 0:
        break;
    else:
        if sheet.cell_value(i, lastCol) != "":
            product = [sheet.cell_value(i, BOARD_NAME), sheet.cell_value(i, lastCol)]
            products.append(product)

for elem in products:
    print(elem[0] + " " +str(elem[1]))
'''for elem in products:
    print("Name: " + elem[0] + "\t\t| " + "Price: " + str(elem[1]))
    f.write(elem[0] + " " + str(elem[1]) + "\n")'''

f.close()
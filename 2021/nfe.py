import xlrd

f = open("items.xml", "w")

# Give the location of the file
loc = ("nfe_teste.xlsx")
 
# To open Workbook
wb = xlrd.open_workbook(loc)

sheet = wb.sheet_by_name("Teste")
print(sheet.cell_value(3,0))

num_rows= sheet.nrows

products = []

for r in range(0, num_rows):
    product = [sheet.cell_value(r, 0), sheet.cell_value(r, 1)]  
    products.append(product)

for elem in products:
    print("Name: " + elem[0] + "\t| " + "Price: " + str(elem[1]))
    f.write(elem[0] + " " + str(elem[1]) + "\n")

f.close()
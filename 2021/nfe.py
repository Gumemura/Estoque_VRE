'''
TO DO

Horario de saida: sempre as 10 da manha. Se emitida antes das 10:00, a data no mesmo dia, caso
    contrario, no ida seguinte

INDUSTRIALIZAÇÃO ou RETORNO DE INDUSTRIALIZAÇÃO: são sempre emitidas duas notas, uma de cada tipo

OK Numero da NFe: ta aqui o desafio

OK Data de vencimento

Minucias de cada item
'''

import xlrd
from datetime import datetime

FIRST_COLUNM = 2
BOARD_CODE = 1
BOARD_NAME = 0

f = open("items.xml", "w")

# Give the location of the file
loc = ("Estoque VRE 2021.xlsx")
 
# To open Workbook
wb = xlrd.open_workbook(loc)
sheet_saida = wb.sheet_by_name("Saida")
sheet_placas = wb.sheet_by_name("Placas")

NF_NUM = int(sheet_saida.cell_value(0, 3))
DATA_VENCIMENTO = datetime(*xlrd.xldate_as_tuple(sheet_saida.cell_value(0, 1), 0)) 

print(NF_NUM)
print(DATA_VENCIMENTO)

num_rows= sheet_saida.nrows

i = FIRST_COLUNM
lastCol = 0
while i < sheet_saida.ncols:
    if sheet_saida.cell_value(1, i) != "":
        lastCol = i
    else:
        break
    i = i + 1

i = 0
products = []
for i in range(2, num_rows):
    if sheet_saida.cell_value(i, BOARD_CODE) == 0:
        break;
    else:
        if sheet_saida.cell_value(i, lastCol) != "":
            # Codicom | Nome | Código | Quantidade | Valor unitário
            product = [sheet_placas.cell_value(i + 1, BOARD_NAME + 2), sheet_saida.cell_value(i, BOARD_NAME), sheet_placas.cell_value(i + 1, BOARD_NAME + 1), sheet_saida.cell_value(i, lastCol), sheet_placas.cell_value(i + 1, BOARD_NAME + 3)]
            products.append(product)

'''for elem in products:
    print("00" + str(int(elem[0])) + " " + elem[1] + " " + str(elem[2]) + " " + str(int(elem[3])) + " " + str(elem[4]) + " " + str(elem[3] * elem[4]))
'''

for elem in products:
    f.write(elem[0] + " " + str(elem[1]) + "\n")

f.close()
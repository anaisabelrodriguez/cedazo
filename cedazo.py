#!/bin/env python
# Cedazo - Morph, Meld and Merge data - Copyright © 2023 Iwan van der Kleijn - See LICENSE.txt for conditions
#from openpyxl import Workbook
import openpyxl

from data import change_colour, remove
#Damos la localización del fichero 
ruta_input = "C:\\Users\\pelviram\\COREPY\\CedazoExcel\\in.xlsx"
ruta_output = "C:\\Users\\pelviram\\COREPY\\CedazoExcel\\out.xlsx"
wb = openpyxl.load_workbook(ruta_input)
ws = wb.active 
#for row in ws: 
remove(ws)
change_colour(ws)
wb.save(ruta_output)



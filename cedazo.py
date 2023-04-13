#!/bin/env python
# Cedazo - Morph, Meld and Merge data - Copyright © 2023 Iwan van der Kleijn - See LICENSE.txt for conditions
#from openpyxl import Workbook
import openpyxl

from data import change_colour, format_column, format_column_iter, format_condition, format_condition_iter, format_condition_iter2, insert_column, remove, unmerge_cells
from miargparse import parser
from openpyxl.styles import PatternFill 
#Damos la localización del fichero de entrada
#ruta_input = "C:\\Users\\lgordoga\\L99PYTHON\\SPRINT_0\\in_corto.xlsx"
#Damos la localización del fichero de salida
#ruta_output = "C:\\Users\\lgordoga\\L99PYTHON\\SPRINT_0\\out.xlsx"
# Creamos objeto wb (libro) de tipo workbook y lo cargamos con lo del excel
args = parser.parse_args()
#ruta_input = args.ruta_input
print(args)

wb = openpyxl.load_workbook(args.ruta_input)
# Creamos objeto ws (hoja), siendo la hoja activa
ws = wb.active

#Metodo que desmergea las celdas de las filas a eliminar
unmerge_cells(ws)

# Metodo que sirve para borrar todas las filas de 1 a 11
#for row in ws: 
remove(ws,1,11)

# Metodo que elimina el color amarillo de todas la celdas que sean amarillas 
change_colour(ws)

#Metodo que inserta una columna nueva detrás de la columna H poniendo la cabecera y copiando el formato de la columna H
insert_column(ws, colNr=9, headerRow=1, headerVal='FTES. Pdtes.')

#Metodo que da formato a la columna que se ha creado usando range
format_column(ws, colNr= 9)
#Metodo que da formato a la columna que se ha creado utilizando iter_cols
#format_column_iter(ws, colNr= 9)

#Metodo que da formato a la columna que se ha creado según criterio  
color_yellow = PatternFill(start_color="FFFFFF99", end_color="FFFFFF99", fill_type = "solid")
#format_condition(ws, colNr= 9, condition="CRITICA", color=color_yellow)
#Metodo que da formato a la columna que se ha creado según criterio utilizando iter_cols y enumerate
#format_condition_iter(ws, colNr= 9, condition="CRITICA", color=color_yellow)
#Metodo que da formato a la columna que se ha creado según criterio utilizando iter_cols y cell
format_condition_iter2(ws, colNr= 9, condition="CRITICA", color=color_yellow)

# Save the workbook to the output file
wb.save(args.ruta_output)



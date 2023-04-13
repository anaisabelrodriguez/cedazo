#!/bin/env python
# Cedazo - Morph, Meld and Merge data - Copyright © 2023 Iwan van der Kleijn - See LICENSE.txt for conditions
#from openpyxl import Workbook
import openpyxl

from data import add_column, change_background, change_colour, remove, unmerge_cells
#Damos la localización del fichero de entrada
ruta_input = "C:\\temp\\in.xlsx"
#Damos la localización del fichero de salida
ruta_output = "C:\\temp\\out.xlsx"

# Creamos objeto wb (libro) de tipo workbook y lo cargamos con lo del excel
wb = openpyxl.load_workbook(ruta_input)
# Creamos objeto ws (hoja), siendo la hoja activa
ws = wb.active

#Metodo que desmergea las celdas de las filas a eliminar
unmerge_cells(ws)

# Metodo que sirve para borrar todas las filas de 1 a 11
remove(ws,1,11)

# Metodo que elimina el color amarillo de todas la celdas que sean amarillas 
change_colour(ws)

#Método que añade una columna en la posición dada
add_column(ws,9,"FTES. Pdtes.")

#Método para cambiar el fondo de la columna I dependiendo de si el valor de F es critica.
change_background(ws)


# Save the workbook to the output file
wb.save(ruta_output)

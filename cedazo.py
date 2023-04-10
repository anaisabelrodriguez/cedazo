#!/bin/env python
# Cedazo - Morph, Meld and Merge data - Copyright © 2023 Iwan van der Kleijn - See LICENSE.txt for conditions
from openpyxl import Workbook
import openpyxl

def remove(sheet):
    filas = sheet.max_row

    for i in range(filas, 0, -1):
        for cell in sheet[i]:
            if cell.fill.start_color.index 


if __name__ == '__main__':
    ruta_input = "C:\Temp\in.xlsx"
    ruta_output = "C:\Temp\out.xlsx"

    wb = openpyxl.load_workbook(ruta_input)

    ws = wb.active

    remove(ws)

    wb.save(ruta_output)

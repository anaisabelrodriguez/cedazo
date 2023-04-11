from openpyxl.styles import PatternFill

def remove_empty(ws): 
#Sirve para borrar todas las filas que estan vacias, (no se utiliza de momento)
    filas = ws.max_row
    for i in range(filas, 0, -1):
        celdas_vacias = all([cell.value is None for cell in ws[i]])
        if celdas_vacias:
            ws.delete_rows(i, 1)

def remove(ws): 
#Sirve para borrar todas las filas de 1 a 11
    ws.delete_rows(1, 11)


def change_colour(ws): 
#Elimina el color amarillo de todas la celdas.
    filas = ws.max_row

    for i_row in range(filas, 0, -1):
       for cell in ws[i_row]:
            cell.fill = PatternFill(start_color="00FFFFFF", end_color="00FFFFFF", fill_type = "solid")
           
def removeFormatting(ws): 
    # ws is not the worksheet name, but the worksheet object 
    filas = ws.max_row
    for i_row in range(1,11): 
        for cell in ws[i_row]: 
            cell.style = 'Normal'

def unmerged_cell(ws):   
    for merge in list(ws.merged_cells):
        ws.unmerge_cells(range_string=str(merge))
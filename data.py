from openpyxl.styles import PatternFill


def remove_empty(ws): 
    '''Metodo que sirve para borrar todas las filas que estan vacias (no se utiliza de momento porque no sirve para borrar
       lad filas que hay que eliminar pero tienen contenido)'''
    filas = ws.max_row
    for i in range(filas, 0, -1):
        celdas_vacias = all([cell.value is None for cell in ws[i]])
        if celdas_vacias:
            ws.delete_rows(i, 1)

def unmerge_cells(ws):
    '''Metodo que desmergea las celdas de las filas a eliminar'''
    # Buscamos las celdas que estan mergeadas
    for merge in list(ws.merged_cells):
        # Separamos esas celdas mergeadas
        ws.unmerge_cells(range_string=str(merge))

def remove(ws,col_ini,col_fin): 
    '''Metodo que sirve para borrar las filas que nos pasan por parámetro'''
    ws.delete_rows(col_ini, col_fin)


def change_colour(ws): 
    '''Metodo que a todas la celdas les pone color blanco (quita el amarillo) '''
    max_row = ws.max_row

    max_column = ws.max_column

    # Hacemos doble bucle para recoger todos los datos de la tabla que empezaba en fila 12 (ahora empezara en la tabla 
    # destino en la fila 1).
    for i_row in range(1, max_row + 1):
       for i_column in range(1, max_column + 1):
           cell = ws.cell(row = i_row, column = i_column)
           # Cambia para todas las celdas el fondo a blanco
           cell.fill = PatternFill(start_color="FFFFFFFF", end_color = None, fill_type = None)

def add_column(ws,col,cabecera):
    ws.insert_cols(col)
    ws.cell(row=1, column=col).value = cabecera

def change_background(ws):
    max_row = ws.max_row
    for i_row in range(1, max_row + 1):
        if ws.cell(row=i_row,column=6).value == 'CRITICA':
            ws.cell(row=i_row,column=9).fill = PatternFill(start_color="FFFFFF99", end_color = "FFFFFF99", fill_type = 'solid')
    
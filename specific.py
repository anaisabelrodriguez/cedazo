from data import change_value, find_string

def assign_value(ws):
    #Asignamos el valor a la columna I(FTES Pdtes.)
    for cell in ws.iter_cols(min_row=2, max_row=ws.max_row-2, min_col=7, max_col=7):
        for cell_condition in cell:
            try:
                cell_drcha = cell_condition.offset(column=1)
                cell_new = float(cell_condition.value)-float(cell_drcha.value)            
            except:
                cell_new = cell_condition.value
            change_value(ws,cell_condition.row,9,'{0:.2f}'.format(cell_new),None,0)

def comprobar_filas(ws):
    #Se comprueban las filas                
    for i_row in range(2, ws.max_row-1):
        try:
            val_critical = ws.cell(i_row,5).value
            val_criticidad = ws.cell(i_row,6).value
            if val_criticidad != None:
                find_str = find_string(val_criticidad.upper())
            else:
                find_str=False
            #print("val_critical ",val_critical)
            if float(ws.cell(i_row,9).value) == float(0):
                change_value(ws,i_row,6,"CUBIERTA","00008000",3)            
            elif float(ws.cell(i_row,9).value) != 0 and  val_critical.upper() == 'YES':
                change_value(ws,i_row,6,"CRITICA","00FF0000",3)
            elif float(ws.cell(i_row,9).value) != 0 and val_critical.upper() == 'NO' and find_str == True:
                change_value(ws,i_row,6,"URGENTE","FFFFFF00",3)
            elif float(ws.cell(i_row,9).value) != 0 and val_critical.upper() == 'NO' and find_str == False:
                change_value(ws,i_row,6,"NORMAL","FFFFFFFF",3) 
        except AttributeError:
            continue

def remove_draft(ws): 
    '''Metodo que sirve para borrar todas las filas que estan vacias (no se utiliza de momento porque no sirve para borrar
       lad filas que hay que eliminar pero tienen contenido)'''
    for col_condition in ws.iter_cols(min_row=1, min_col=3, max_col=3): 
        #print(col_condition)
        #el siguiente for itera a través de cada celda en la columna seleccionada
        for cell_condition in col_condition:
        #Si el valor de la celda coincide con la condición, entonces se borra 
            value_celda = cell_condition.value
            try:
                if value_celda.upper() == 'DRAFT':
                    ws.delete_rows(cell_condition.row, 1)
            except:
                continue

            
            
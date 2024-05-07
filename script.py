from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from interface import *

def activate(titulo, name):
    
    wb = load_workbook(file)

    ws = wb.active

    if name != "Municipais - Ocultar":
        ws.insert_rows(idx=1, amount=1)
    
        ws['A1'] = titulo

        cell = ws['A1']
        
        cell.font = Font(name='Calibri',
                        size='18',
                        bold=True)
        
        ws.row_dimensions[1].height = 30

        for col in range(1,30):
            cell = ws.cell(2,col)
            cell.font = Font(bold=True)
    redFill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
    # Hapvida - Colunas que não são hidden: E F P R S U V + W  *** K e M se tiver dengue
    if name == "Hapvida":
        for c in ('B', 'C', 'D', 'G', 'H', 'I', 'J', 'L', 'N', 'O', 'Q', 'R', 'T'):
            ws.column_dimensions[c].hidden= True
            color = c + '2'
            ws[color].fill = redFill
        ws.merge_cells('A1:X1')
        ws['X2'] = 'Dist'
        ws['F2'] = 'DataNasc'
        ws['A2'] = 'DataSis'
        ws['P2'] = 'Proced'
        ws['S2'] = 'Result'

    # NS1 - Colunas que não são hidden: D E F H L R S T U W X
    elif name == "NS1":

        # for c in ('A', 'B', 'C', 'G', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q', 'V','Z'):
        # Abaixo: Utilizada APENAS quando a sorologia está suspensa! Se há sorologia utilizar a de cima.
        for c in ('A', 'B', 'C', 'G', 'I', 'J', 'K', 'M', 'N', 'O', 'P', 'Q','U', 'V','W', 'X', 'Z'):
            ws.column_dimensions[c].hidden= True
            color = c + '2'
            ws[color].fill = redFill
        ws.merge_cells('A1:Y1')
        ws['H2'] = 'DataNasc'
        ws['A2'] = 'DataSis'
        ws['X2'] = 'Num'
        ws['P2'] = 'Proced'
    
    elif name == "GAL":
        for c in ('C','D','I','J','L','N'):
            ws.column_dimensions[c].hidden= True
        ws.merge_cells('A1:Q1')
    
    elif name == "Sabin":
        ws = wb['POSITIVO']

        ws.insert_rows(idx=1, amount=1)
    
        ws['A1'] = titulo

        cell = ws['A1']
        
        cell.font = Font(name='Calibri',
                        size='18',
                        bold=True)
        
        ws.row_dimensions[1].height = 30

        for col in range(1,30):
            cell = ws.cell(2,col)
            cell.font = Font(bold=True)
        for c in ('A', 'F', 'G', 'J', 'K', 'L', 'M', 'N'):
            ws.column_dimensions[c].hidden= True
        ws.merge_cells('A1:Q1')
        ws['E2'] = "Dist"
        ws['I2'] = "DataNasc"

    elif name == "Municipais - Ocultar":
        for c in ('A', 'B', 'C', 'E', 'H', 'J', 'K', 'P'):
            ws.column_dimensions[c].hidden= True                

    for row in range(1,2500):
        for col in range(1,80):
            cell = ws.cell(row,col)
            cell.alignment = Alignment(horizontal='center', vertical='center')

    wb.save(file)

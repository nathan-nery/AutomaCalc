from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from interface import *

# Salvar arquivo após fazer a alteração:
# workbook.save('C:\\Users\\saude\\Desktop\\Nathan\\Projetos\\EditorDeCalc\\teste.xlsx')

def activate(titulo, name):

    #local = input('Qual o diretório do arquivo? ')
    
    wb = load_workbook(file)

    ws = wb.active

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
    
    # Hapvida - Colunas que não são hidden: E F P R S U V + W  *** K e M se tiver dengue
    if name == "Hapvida":
        for c in ('A', 'B', 'C', 'D', 'G', 'H', 'I', 'J', 'L', 'N', 'O', 'Q', 'T'):
            ws.column_dimensions[c].hidden= True
        ws.merge_cells('A1:W1')

# NS1 - Colunas que não são hidden: D E F H L M R S T U W X
    elif name == "NS1":
        for d in ('A', 'B', 'C', 'G', 'I', 'J', 'K', 'N', 'O', 'P', 'Q', 'V'):
            ws.column_dimensions[d].hidden= True
        ws.merge_cells('A1:X1')

    for row in range(1,80):
        for col in range(1,80):
            cell = ws.cell(row,col)
            cell.alignment = Alignment(horizontal='center', vertical='center')

    wb.save(file)

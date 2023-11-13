from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from interface import *

wb = Workbook()

ws = wb.active
# Salvar arquivo após fazer a alteração:
# workbook.save('C:\\Users\\saude\\Desktop\\Nathan\\Projetos\\EditorDeCalc\\teste.xlsx')

def activate(titulo):

    #local = input('Qual o diretório do arquivo? ')
    
    load_workbook(file)

    ws['A1'] = 0

    ws.insert_rows(idx=1, amount=1)

    ws.merge_cells('A1:W1')

    ws['A1'] = titulo

    cell = ws['A1']
    cell.font = Font(name='Calibri',
                    size='18',
                    bold=True)
    ws.row_dimensions[1].height = 30

    for col in range(1,30):
        cell = ws.cell(2,col)
        cell.font = Font(bold=True)

    for row in range(1,60):
        for col in range(1,60):
            cell = ws.cell(row,col)
            cell.alignment = Alignment(horizontal='center')

    wb.save(file)

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
import openpyxl
from interface import *

workbook = Workbook()

worksheet = workbook.active
# Salvar arquivo após fazer a alteração:
# workbook.save('C:\\Users\\saude\\Desktop\\Nathan\\Projetos\\EditorDeCalc\\teste.xlsx')

def activate():

# local = input('Qual o diretório do arquivo? ')
    #wb = openpyxl.load_workbook(local)

    worksheet['A1'] = 0

    worksheet.insert_rows(idx=1, amount=1)

    worksheet.merge_cells('A1:W1')

    worksheet['A1'] = titulo

    cell = worksheet['A1']

    cell.font = Font(name='Calibri',
                    size='18',
                    bold=True)

    worksheet.row_dimensions[1].height = 30
    for row in range(1,60):
        for col in range(1,60):
            cell = worksheet.cell(row,col)
            cell.alignment = Alignment(horizontal='center')

    #workbook.save(local)
# definir local por meio de janela que abre e acessa os arquivos para coletar diretório

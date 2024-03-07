import json
import xlsxwriter

# Carrega o arquivo JSON
with open("teste.json", "r") as f:
    data = json.load(f)

# Cria a pasta de trabalho e a planilha
workbook = xlsxwriter.Workbook("planilha.xlsx")
worksheet = workbook.add_worksheet()

# Define a formatação para a coluna de recuo
indent_format = workbook.add_format()
indent_format.set_indent(1)

row = 0

def processa_dict(item, prev_col, headers: bool = False):
    global row
    col = prev_col
    if headers:
        for key, value in item.items():
            worksheet.write(row, col, key, indent_format)
            col += 1
        row += 1

    col = prev_col
    for key, value in item.items():
        if type(value) == list:
            row += 1
            processa_lista(value, col)
        else:
            worksheet.write(row, col, value, indent_format)
        col += 1


def processa_lista(item, prev_col):
    global row
    headers = True
    for items in item:
        tipo = type(items)
        if tipo == dict:
            processa_dict(items, prev_col, headers)
            headers = False
        elif tipo == list:
            processa_lista(items, prev_col+1)
        row += 1

if type(data) == list:
    processa_lista(data, 0)
else:
    processa_dict(data, 0)

workbook.close()
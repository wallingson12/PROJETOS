import os
import xlrd
from openpyxl import Workbook

pasta = r"C:\Users\wallingson.silva\TO DO\UNIFICAR XLS PARA PDF\DGCAs 2020_2022"
novo_arquivo = r"C:\Users\wallingson.silva\TO DO\UNIFICAR XLS PARA PDF\DGCAs 2020_2022\arquivo.xlsx"

novo_workbook = Workbook()
novo_planilha = novo_workbook.active

# Nome da sheet
novo_planilha.title = "Valores O31"

# Define os nomes das colunas
novo_planilha.append(["Data", "Valor acumulado"])

for arquivo in os.listdir(pasta):
    if arquivo.endswith(".xls"):
        caminho_arquivo = os.path.join(pasta, arquivo)
        planilha = xlrd.open_workbook(caminho_arquivo).sheet_by_index(0)

        # Selecionando as linhas requisitadas
        A_15 = planilha.cell_value(14, 0)  # Célula A15 (linha 15, coluna 1)
        O_31 = planilha.cell_value(30, 14)  # Célula O31 (linha 31, coluna 15)

        # Gerando o arquivo
        novo_planilha.append([A_15, O_31])

# Salva o novo arquivo Excel
novo_workbook.save(novo_arquivo)
import os
import comtypes.client as client


def convert_xls_to_pdf(directory, output_file, file_sequence):
    excel = client.CreateObject("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    # armazena o nome do arquivo
    pdf_filename = output_file

    # Cria um novo Workbook para armazenar as planilhas dos arquivos Excel
    merged_workbook = excel.Workbooks.Add()

    # Para cada dado na lista file_sequence
    for file_number in reversed(file_sequence):
        xls_file = f"{file_number}.xls"
        full_path = os.path.join(directory, xls_file)
        workbook = excel.Workbooks.Open(full_path)

        # Copia as planilhas do arquivo Excel para o Workbook unificado
        workbook.Sheets.Copy(Before=merged_workbook.Sheets(1))

        workbook.Close()

    # Salva o Workbook unificado como um arquivo PDF
    merged_workbook.ExportAsFixedFormat(0, pdf_filename, 0)

    # Fecha o Workbook unificado
    merged_workbook.Close(False)

    # Renomeia o arquivo PDF resultante para o nome desejado
    os.rename(pdf_filename, output_file)

    excel.Quit()


# Diretório contendo os arquivos Excel
directory = r"C:\Users\wallingson.silva\TO DO\UNIFICAR XLS PARA PDF\DGCAs 2020_2022"

# Caminho e nome do arquivo PDF de saída
output_file = r"C:\Users\wallingson.silva\TO DO\UNIFICAR XLS PARA PDF\DGCAs 2020_2022\arquivo.pdf"

# Sequência personalizada dos arquivos
file_sequence = [
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 012020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 022020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 032020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 042020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 052020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 062020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 072020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 082020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 092020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 102020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 112020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 122020",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 012021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 022021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 032021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 042021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 052021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 062021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 072021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 082021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 092021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 102021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 112021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 122021",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 012022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 022022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 032022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 042022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 052022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 062022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 072022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 082022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 092022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 102022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 112022",
    "DGCA - NILIT AMERICANA FIBRAS DE POLIAMIDA LTDA - 122022"
]

# Chama a função para converter os arquivos Excel para PDF
convert_xls_to_pdf(directory, output_file, file_sequence)

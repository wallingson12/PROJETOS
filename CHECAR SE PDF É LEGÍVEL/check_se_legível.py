import pdfminer.high_level as pdf
import os

# Diretório com os arquivos PDF
diretorio = r"C:\Users\wallingson.silva\Downloads\PROJETO\Projeto Macan\DCTFs\3.1.3 - DCTF"

# Lista para armazenar os caminhos dos arquivos PDF
caminhos_pdf = []

# Itera sobre todos os arquivos no diretório
for nome_arquivo in os.listdir(diretorio):
    if nome_arquivo.endswith(".pdf"):  # Verifica se é um arquivo PDF
        caminho_arquivo = os.path.join(diretorio, nome_arquivo)
        caminhos_pdf.append(caminho_arquivo)

# Contador de arquivos processados com sucesso
arquivos_ok = 0

# Itera sobre os caminhos dos arquivos PDF
for caminho in caminhos_pdf:
    # Extrai o texto do PDF
    text = pdf.extract_text(caminho)

    if text is not None and text.strip():  # Verifica se o texto não é vazio
        # Extrai as informações do texto
        lines = text.split('\n')
        for i in range(len(lines)):
            # Printa as linhas
            print(f"Line {i + 1}: {lines[i]}")

            # if lines[i].startswith('IRPJ'):
            # if i + 1 < len(lines):
            # found_text = lines[i + 1].strip()
            # print(found_text)
            # break

        # Se o texto foi extraído com sucesso, incrementa o contador
        arquivos_ok += 1
    else:
        print(f"Não foi possível extrair texto do arquivo: {caminho}")

# Mostra o número de arquivos OK
print(f"\nNúmero de arquivos processados com sucesso: {arquivos_ok}")

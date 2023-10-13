import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk

arquivo_selecionado = ""

def processar_planilha():
    coluna_divisao = input_coluna.get()
    if not arquivo_selecionado:
        messagebox.showwarning("Aviso", "Por favor, selecione um arquivo Excel.")
        return

    try:
        # Determine a planilha com os dados
        df = pd.read_excel(arquivo_selecionado)

        # Obtemos os valores únicos da coluna de agrupamento
        valores_unicos = df[coluna_divisao].unique()

        # Especifique o diretório onde os arquivos de saída serão salvos
        diretorio = os.getcwd()

        # Cria uma pasta chamada "output" dentro do diretório atual, se ela ainda não existir
        diretorio_output = os.path.join(diretorio, "output")
        if not os.path.exists(diretorio_output):
            os.makedirs(diretorio_output)

        # Loop para criar um arquivo Excel para cada valor único
        for valor in valores_unicos:
            # Filtra o DataFrame com o valor único da coluna 'Coluna'
            novo_df = df[df[coluna_divisao] == valor]

            # Define o nome do arquivo a ser gerado "nome_arquivo_valor_unico.xlsx"
            nome_arquivo_saida = f"{valor}.xlsx"

            # Cria o novo arquivo Excel com o nome estabelecido no diretório de saída
            caminho_saida = os.path.join(diretorio_output, nome_arquivo_saida)
            novo_df.to_excel(caminho_saida, index=False)

        messagebox.showinfo("Concluído", "Processamento concluído com sucesso! Os arquivos foram salvos na pasta 'output'.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento: {str(e)}")

def selecionar_arquivo():
    global arquivo_selecionado
    arquivo_selecionado = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx")])
    if arquivo_selecionado:
        messagebox.showinfo("Sucesso", f"Arquivo selecionado: {arquivo_selecionado}")

# Criação da interface gráfica
root = tk.Tk()
root.title("Divisor de Planilhas")

path_to_icon = "MCS.ico"
if os.path.exists(path_to_icon):
    root.iconbitmap(default=path_to_icon)

largura_display = 600
altura_display = 400

largura_tela = root.winfo_screenwidth()
altura_tela = root.winfo_screenheight()

posicao_x = (largura_tela - largura_display) // 2
posicao_y = (altura_tela - altura_display) // 2

root.geometry(f"{largura_display}x{altura_display}+{posicao_x}+{posicao_y}")

root.configure(bg="blue")

path_to_image = "MCS.png"
if os.path.exists(path_to_image):
    image = Image.open(path_to_image)
    image.thumbnail((200, 200))  # Redimensiona a imagem para caber no botão
    photo = ImageTk.PhotoImage(image)
else:
    photo = None

frame_central = tk.Frame(root, bg="#00426b")
frame_central.pack(fill=tk.BOTH, expand=True)

if photo:
    image_label = tk.Label(frame_central, image=photo, bg="#00426b")
    image_label.pack(pady=(20, 0))

label_arquivo = tk.Label(frame_central, text="Selecione o arquivo Excel:", bg="#00426b", fg="white")
label_arquivo.pack(pady=5)

selecionar_arquivo_button = tk.Button(frame_central, text="Selecionar arquivo", command=selecionar_arquivo)
selecionar_arquivo_button.pack(pady=5)

label_coluna = tk.Label(frame_central, text="Insira o nome da coluna para divisão:", bg="#00426b", fg="white")
label_coluna.pack(pady=5)

input_coluna = tk.Entry(frame_central, width=30)
input_coluna.pack(pady=5)

processar_button = tk.Button(frame_central, text="Processar", command=processar_planilha)
processar_button.pack(pady=15)

root.mainloop()

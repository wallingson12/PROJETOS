import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import camelot

arquivo_selecionado = ""

def extrair_tabelas_para_excel(pdf_file):
    tabelas = camelot.read_pdf(pdf_file, pages='all', flavor='stream')

    nome_arquivo = os.path.splitext(os.path.basename(pdf_file))[0]
    df_final = pd.DataFrame()

    for tabela in tabelas:
        df_final = pd.concat([df_final, tabela.df], ignore_index=False)

    output_dir = os.path.join(os.path.dirname(pdf_file), nome_arquivo)
    os.makedirs(output_dir, exist_ok=True)

    output_file = os.path.join(output_dir, f'{nome_arquivo}.xlsx')
    df_final.to_excel(output_file, header=False, index=False)

def processar_pdfs():
    global pasta_origem
    if not pasta_origem:
        messagebox.showwarning("Aviso", "Por favor, selecione uma pasta de origem.")
        return

    for root, dirs, files in os.walk(pasta_origem):
        for file in files:
            if file.endswith(".pdf"):
                pdf_file = os.path.join(root, file)
                extrair_tabelas_para_excel(pdf_file)

def selecionar_pasta_saida():
    global pasta_saida
    pasta_saida = filedialog.askdirectory(title="Selecione a pasta de saída")
    if pasta_saida:
        messagebox.showinfo("Sucesso", f"Pasta de saída selecionada: {pasta_saida}")

def selecionar_pasta_origem():
    global pasta_origem
    pasta_origem = filedialog.askdirectory(title="Selecione a pasta de origem")
    if pasta_origem:
        messagebox.showinfo("Sucesso", f"Pasta de origem selecionada: {pasta_origem}")

root = tk.Tk()
root.title("Divisor de Planilhas")

path_to_icon = "icone.png"
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

path_to_image = "icone.png"
if os.path.exists(path_to_image):
    image = Image.open(path_to_image)
    image.thumbnail((200, 200))
    photo = ImageTk.PhotoImage(image)
else:
    photo = None

frame_central = tk.Frame(root, bg="#00426b")
frame_central.pack(fill=tk.BOTH, expand=True)

if photo:
    image_label = tk.Label(frame_central, image=photo, bg="#00426b")
    image_label.pack(pady=(20, 0))

label_pasta_origem = tk.Label(frame_central, text="Selecione a pasta de origem:", bg="#00426b", fg="white")
label_pasta_origem.pack(pady=5)

selecionar_pasta_origem_button = tk.Button(frame_central, text="Selecionar pasta", command=selecionar_pasta_origem)
selecionar_pasta_origem_button.pack(pady=5)

label_pasta_saida = tk.Label(frame_central, text="Selecione a pasta de saída:", bg="#00426b", fg="white")
label_pasta_saida.pack(pady=5)

selecionar_pasta_saida_button = tk.Button(frame_central, text="Selecionar pasta", command=selecionar_pasta_saida)
selecionar_pasta_saida_button.pack(pady=5)

processar_button = tk.Button(frame_central, text="Processar PDFs", command=processar_pdfs)
processar_button.pack(pady=15)

root.mainloop()

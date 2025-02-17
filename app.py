import xlwings as xw
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from pyautocad import Autocad, APoint  # Biblioteca para lidar com arquivos DWG


# Definições de cálculo
PESO_CHAPA_KG_M2 = 7.85  # Exemplo de densidade da chapa
PRECO_KG = 15.00  # Preço por kg da chapa


def ler_dwg(arquivo_dwg):
    """
    Lê o arquivo DWG e extrai as dimensões dos dutos.
    Retorna uma lista de dicionários com largura, altura, comprimento e peso estimado.
    """
    try:
        acad = Autocad(create_if_not_exists=True)
        acad.open(arquivo_dwg)
       
        materiais = []
       
        # Itera sobre as polilinhas (LWPOLYLINE) no arquivo DWG
        for entidade in acad.iter_objects('LWPOLYLINE'):
            if len(entidade) >= 4:  # Certifica-se de que há 4 pontos na polilinha
                pontos = entidade.get_points()
                largura = abs(pontos[1][0] - pontos[0][0])
                altura = abs(pontos[2][1] - pontos[1][1])
                comprimento = entidade.elevation  # Usado como comprimento
                area_m2 = (largura / 1000) * (altura / 1000) * comprimento
                peso_kg = area_m2 * PESO_CHAPA_KG_M2
                preco_total = peso_kg * PRECO_KG


                materiais.append({
                    "Largura (cm)": largura,
                    "Altura (cm)": altura,
                    "Comprimento (m)": comprimento,
                    "Área (m²)": area_m2,
                    "Peso (kg)": peso_kg,
                    "Preço (R$)": preco_total
                })
        return materiais
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao ler o arquivo DWG.\nErro: {e}")
        return []


def atualizar_planilha(materiais, caminho_excel):
    """
    Atualiza a planilha do Excel com os materiais extraídos do DWG.
    """
    if not os.path.exists(caminho_excel):
        messagebox.showerror("Erro", "Arquivo Excel não encontrado.")
        return


    try:
        app = xw.App(visible=False)  
        wb = xw.Book(caminho_excel)
        ws = wb.sheets["Orçamento"]  
        ws.range("A2:F100").clear_contents()
        ws.range("A1").value = ["Largura (cm)", "Altura (cm)", "Comprimento (m)", "Área (m²)", "Peso (kg)", "Preço (R$)"]
        ws.range("A2").value = [list(mat.values()) for mat in materiais]
        wb.save()
        wb.close()
        app.quit()
        messagebox.showinfo("Sucesso", "Planilha atualizada com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao atualizar a planilha.\nErro: {e}")


def selecionar_arquivo_dwg():
    """
    Abre uma janela para o usuário selecionar o arquivo DWG.
    """
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos AutoCAD", "*.dwg")])
    if arquivo:
        entry_dwg.delete(0, tk.END)
        entry_dwg.insert(0, arquivo)


def selecionar_arquivo_excel():
    """
    Abre uma janela para o usuário selecionar o arquivo Excel.
    """
    arquivo = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")])
    if arquivo:
        entry_excel.delete(0, tk.END)
        entry_excel.insert(0, arquivo)


def executar():
    """
    Executa o fluxo completo ao clicar no botão.
    """
    arquivo_dwg = entry_dwg.get()
    caminho_excel = entry_excel.get()


    if not os.path.exists(arquivo_dwg):
        messagebox.showerror("Erro", "Selecione um arquivo DWG válido.")
        return
    if not os.path.exists(caminho_excel):
        messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
        return


    materiais = ler_dwg(arquivo_dwg)


    if not materiais:
        messagebox.showwarning("Aviso", "Nenhum material encontrado no DWG.")
        return


    atualizar_planilha(materiais, caminho_excel)


# Criar a interface gráfica
root = tk.Tk()
root.title("Automação de Orçamentos - Refrigeração")
root.geometry("500x250")


# Rótulos e campos de entrada
tk.Label(root, text="Selecione o arquivo DWG:").pack(pady=5)
entry_dwg = tk.Entry(root, width=50)
entry_dwg.pack()
tk.Button(root, text="Procurar", command=selecionar_arquivo_dwg).pack(pady=5)


tk.Label(root, text="Selecione o arquivo Excel:").pack(pady=5)
entry_excel = tk.Entry(root, width=50)
entry_excel.pack()
tk.Button(root, text="Procurar", command=selecionar_arquivo_excel).pack(pady=5)


# Botão para executar
tk.Button(root, text="Executar", command=executar, bg="green", fg="white").pack(pady=20)


root.mainloop()




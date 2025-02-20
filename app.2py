import xlwings as xw
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import ezdxf  # Biblioteca para lidar com arquivos DXF

#Tj7

# Definições de cálculo
PESO_CHAPA_KG_M2 = 7.85  # Exemplo de densidade da chapa
PRECO_KG = 15.00  # Preço por kg da chapa


def ler_dxf(arquivo_dxf):
    """
    Lê o arquivo DXF e extrai as dimensões dos dutos.
    Retorna uma lista de dicionários com largura, altura, comprimento e peso estimado.
    """
    try:
        doc = ezdxf.readfile(arquivo_dxf)
        msp = doc.modelspace()


        materiais = []


        # Itera sobre as entidades no arquivo DXF
        for entidade in msp:
            if entidade.dxftype() == 'LWPOLYLINE':
                pontos = entidade.get_points('xy')
                if len(pontos) >= 4:  # Certifica-se de que há 4 pontos na polilinha
                    largura = abs(pontos[1][0] - pontos[0][0])
                    altura = abs(pontos[2][1] - pontos[1][1])
                    comprimento = entidade.dxf.elevation  # Usado como comprimento
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
        messagebox.showerror("Erro", f"Falha ao ler o arquivo DXF.\nErro: {e}")
        return []


def atualizar_planilha(materiais, caminho_excel):
    """
    Atualiza a planilha do Excel com os materiais extraídos do DXF.
    """
    if not os.path.exists(caminho_excel):
        messagebox.showerror("Erro", "Arquivo Excel não encontrado.")
        return


    try:
        app = xw.App(visible=False)  
        wb = xw.Book(caminho_excel)


        # Verifica se a planilha "Orçamento" existe
        if "Orçamento" not in [sheet.name for sheet in wb.sheets]:
            messagebox.showerror("Erro", 'A planilha "Orçamento" não foi encontrada no arquivo Excel.')
            wb.close()
            app.quit()
            return


        ws = wb.sheets["Orçamento"]


        # Limpa o conteúdo da planilha
        ws.range("A2:F100").clear_contents()


        # Insere os cabeçalhos
        ws.range("A1").value = ["Largura (cm)", "Altura (cm)", "Comprimento (m)", "Área (m²)", "Peso (kg)", "Preço (R$)"]


        # Insere os dados
        if materiais:  # Verifica se há dados para inserir
            dados = [list(mat.values()) for mat in materiais]
            ws.range("A2").value = dados


        wb.save()
        wb.close()
        app.quit()
        messagebox.showinfo("Sucesso", "Planilha atualizada com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao atualizar a planilha.\nErro: {e}")
        if 'wb' in locals():  # Fecha o workbook se ele estiver aberto
            wb.close()
        if 'app' in locals():  # Fecha o aplicativo do Excel se estiver aberto
            app.quit()


def selecionar_arquivo_dxf():
    """
    Abre uma janela para o usuário selecionar o arquivo DXF.
    """
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos DXF", "*.dxf")])
    if arquivo:
        entry_dxf.delete(0, tk.END)
        entry_dxf.insert(0, arquivo)


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
    arquivo_dxf = entry_dxf.get()
    caminho_excel = entry_excel.get()


    if not os.path.exists(arquivo_dxf):
        messagebox.showerror("Erro", "Selecione um arquivo DXF válido.")
        return
    if not os.path.exists(caminho_excel):
        messagebox.showerror("Erro", "Selecione um arquivo Excel válido.")
        return


    materiais = ler_dxf(arquivo_dxf)


    if not materiais:
        messagebox.showwarning("Aviso", "Nenhum material encontrado no DXF.")
        return


    atualizar_planilha(materiais, caminho_excel)


# Criar a interface gráfica
root = tk.Tk()
root.title("Automação de Orçamentos - Refrigeração")
root.geometry("500x250")


# Rótulos e campos de entrada
tk.Label(root, text="Selecione o arquivo DXF:").pack(pady=5)
entry_dxf = tk.Entry(root, width=50)
entry_dxf.pack()
tk.Button(root, text="Procurar", command=selecionar_arquivo_dxf).pack(pady=5)


tk.Label(root, text="Selecione o arquivo Excel:").pack(pady=5)
entry_excel = tk.Entry(root, width=50)
entry_excel.pack()
tk.Button(root, text="Procurar", command=selecionar_arquivo_excel).pack(pady=5)


# Botão para executar
tk.Button(root, text="Executar", command=executar, bg="green", fg="white").pack(pady=20)


root.mainloop()
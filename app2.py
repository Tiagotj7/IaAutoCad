import xlwings as xw
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import ezdxf


PESO_CHAPA = {
    0.61: 4.88, 0.68: 5.49, 0.76: 6.1, 0.84: 6.71, 0.91: 7.32, 1.06: 8.54,
    1.21: 9.76, 1.37: 10.98, 1.52: 12.21, 1.71: 13.73, 1.9: 15.26, 2.28: 18.81,
    2.66: 21.36, 3.04: 24.41, 3.42: 27.46, 3.8: 30.52, 4.18: 33.57, 4.55: 36.62,
    4.76: 37.35, 6.35: 49.8, 7.94: 62.25, 9.53: 74.7, 12.7: 99.6, 15.88: 124.49,
    19.05: 149.39, 22.23: 174.29, 25.4: 199.19, 26.99: 211.64, 28.58: 224.09,
    30.16: 236.53, 31.75: 249.98, 33.34: 261.43, 34.93: 273.88
}
PRECO_KG = 15.00


def ler_dxf(arquivo_dxf):
    try:
        doc = ezdxf.readfile(arquivo_dxf)
        msp = doc.modelspace()
        materiais = []
        for entidade in msp:
            if entidade.dxftype() == 'LWPOLYLINE':
                pontos = entidade.get_points('xy')
                if len(pontos) >= 4:
                    largura = abs(pontos[1][0] - pontos[0][0])
                    altura = abs(pontos[2][1] - pontos[1][1])
                    comprimento = entidade.dxf.elevation
                    espessura = min(PESO_CHAPA.keys(), key=lambda x: abs(x - altura))
                    peso_kg = (largura / 1000) * (altura / 1000) * comprimento * PESO_CHAPA[espessura]
                    preco_total = peso_kg * PRECO_KG
                    materiais.append({
                        "Largura (cm)": round(largura, 2),
                        "Altura (cm)": round(altura, 2),
                        "Comprimento (m)": round(comprimento, 2),
                        "Espessura (mm)": round(espessura, 2),
                        "Peso (kg)": round(peso_kg, 2),
                        "Preço (R$)": round(preco_total, 2)
                    })
        return materiais
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao ler o arquivo DXF.\nErro: {e}")
        return []


def atualizar_planilha(materiais, caminho_excel):
    if not os.path.exists(caminho_excel):
        messagebox.showerror("Erro", "Arquivo Excel não encontrado.")
        return
    try:
        app = xw.App(visible=False)
        wb = xw.Book(caminho_excel)
        if "Orçamento" not in [sheet.name for sheet in wb.sheets]:
            messagebox.showerror("Erro", 'A planilha "Orçamento" não foi encontrada no arquivo Excel.')
            wb.close()
            app.quit()
            return
        ws = wb.sheets["Orçamento"]
        ws.range("A2:G100").clear_contents()
        ws.range("A1").value = ["Largura (cm)", "Altura (cm)", "Comprimento (m)", "Espessura (mm)", "Peso (kg)", "Preço (R$)"]
        if materiais:
            dados = [list(mat.values()) for mat in materiais]
            ws.range("A2").value = dados
            ws.range(f"E2:E{len(dados) + 1}").number_format = "0.00"
            ws.range(f"F2:F{len(dados) + 1}").number_format = "R$ #,##0.00"
        wb.save()
        wb.close()
        app.quit()
        messagebox.showinfo("Sucesso", "Planilha atualizada com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao atualizar a planilha.\nErro: {e}")
        if 'wb' in locals():
            wb.close()
        if 'app' in locals():
            app.quit()


def selecionar_arquivo_dxf():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos DXF", "*.dxf")])
    if arquivo:
        entry_dxf.delete(0, tk.END)
        entry_dxf.insert(0, arquivo)


def selecionar_arquivo_excel():
    arquivo = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")])
    if arquivo:
        entry_excel.delete(0, tk.END)
        entry_excel.insert(0, arquivo)


def executar():
    arquivo_dxf = entry_dxf.get()
    caminho_excel = entry_excel.get()
    if not os.path.exists(arquivo_dxf) or not os.path.exists(caminho_excel):
        messagebox.showerror("Erro", "Selecione arquivos válidos.")
        return
    materiais = ler_dxf(arquivo_dxf)
    if not materiais:
        messagebox.showwarning("Aviso", "Nenhum material encontrado no DXF.")
        return
    atualizar_planilha(materiais, caminho_excel)
root = tk.Tk()
root.title("Automação de Orçamentos - Refrigeração")
root.geometry("500x250")
tk.Label(root, text="Selecione o arquivo DXF:").pack(pady=5)
entry_dxf = tk.Entry(root, width=50)
entry_dxf.pack()
tk.Button(root, text="Procurar", command=selecionar_arquivo_dxf).pack(pady=5)
tk.Label(root, text="Selecione o arquivo Excel:").pack(pady=5)
entry_excel = tk.Entry(root, width=50)
entry_excel.pack()
tk.Button(root, text="Procurar", command=selecionar_arquivo_excel).pack(pady=5)
tk.Button(root, text="Executar", command=executar, bg="green", fg="white").pack(pady=20)
root.mainloop()




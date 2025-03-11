import xlwings as xw
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ezdxf

PESO_CHAPA = {0.61: 4.88, 0.68: 5.49, 0.76: 6.1, 0.84: 6.71, 0.91: 7.32, 1.06: 8.54,
    1.21: 9.76, 1.37: 10.98, 1.52: 12.21, 1.71: 13.73, 1.9: 15.26, 2.28: 18.81,
    2.66: 21.36, 3.04: 24.41, 3.42: 27.46, 3.8: 30.52, 4.18: 33.57, 4.55: 36.62,
    4.76: 37.35, 6.35: 49.8, 7.94: 62.25, 9.53: 74.7, 12.7: 99.6, 15.88: 124.49,
    19.05: 149.39, 22.23: 174.29, 25.4: 199.19, 26.99: 211.64, 28.58: 224.09,
    30.16: 236.53, 31.75: 249.98, 33.34: 261.43, 34.93: 273.88}
PRECO_KG = 15.00

def adicionar_peca():
    if entry_largura.get() and entry_altura.get() and entry_comprimento.get() and entry_espessura.get():
        tree.insert("", "end", values=(entry_largura.get(), entry_altura.get(), entry_comprimento.get(), entry_espessura.get()))
        limpar_campos()

def remover_peca():
    item_selecionado = tree.selection()
    if item_selecionado:
        tree.delete(item_selecionado)

def editar_peca():
    item_selecionado = tree.selection()
    if item_selecionado:
        valores = tree.item(item_selecionado, "values")
        entry_largura.delete(0, tk.END)
        entry_altura.delete(0, tk.END)
        entry_comprimento.delete(0, tk.END)
        entry_espessura.delete(0, tk.END)
        entry_largura.insert(0, valores[0])
        entry_altura.insert(0, valores[1])
        entry_comprimento.insert(0, valores[2])
        entry_espessura.insert(0, valores[3])
        tree.delete(item_selecionado)

def calcular_manual():
    total_peso = 0
    total_preco = 0
    for item in tree.get_children():
        valores = tree.item(item, "values")
        try:
            largura, altura, comprimento, espessura = map(float, valores)
            if espessura not in PESO_CHAPA:
                messagebox.showerror("Erro", "Espessura não encontrada na tabela.")
                return
            peso_kg = (largura / 1000) * (altura / 1000) * comprimento * PESO_CHAPA[espessura]
            preco_total = round(peso_kg * PRECO_KG, 2)
            total_peso += peso_kg
            total_preco += preco_total
        except ValueError:
            messagebox.showerror("Erro", "Erro nos valores inseridos.")
            return
    label_resultado.config(text=f"Peso Total: {total_peso:.2f} kg | Preço Total: R$ {total_preco:.2f}")

def limpar_campos():
    entry_largura.delete(0, tk.END)
    entry_altura.delete(0, tk.END)
    entry_comprimento.delete(0, tk.END)
    entry_espessura.delete(0, tk.END)

def alternar_frame():
    frame_manual.pack() if var_opcao.get() else frame_manual.pack_forget()

root = tk.Tk()
root.title("Automação de Orçamentos")
root.geometry("800x600")

frame_arquivos = tk.Frame(root)
frame_arquivos.pack()
tk.Label(frame_arquivos, text="Arquivo DXF:").pack()
entry_dxf = tk.Entry(frame_arquivos, width=50)
entry_dxf.pack()
tk.Button(frame_arquivos, text="Procurar", command=lambda: entry_dxf.insert(0, filedialog.askopenfilename(filetypes=[("DXF", "*.dxf")]))).pack()
tk.Label(frame_arquivos, text="Arquivo Excel:").pack()
entry_excel = tk.Entry(frame_arquivos, width=50)
entry_excel.pack()
tk.Button(frame_arquivos, text="Procurar", command=lambda: entry_excel.insert(0, filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")]))).pack()
tk.Button(frame_arquivos, text="Executar", bg="green", fg="white").pack(pady=10)

var_opcao = tk.BooleanVar()
tk.Checkbutton(root, text="Opção Personalizada", variable=var_opcao, command=alternar_frame).pack()

frame_manual = tk.Frame(root)
frame_manual.pack()

frame_topo = tk.Frame(frame_manual)
frame_topo.pack()
tk.Button(frame_topo, text="Editar Peça", command=editar_peca, bg="yellow", fg="black").pack(side=tk.LEFT, padx=5)
tk.Button(frame_topo, text="Remover Peça", command=remover_peca, bg="red", fg="white").pack(side=tk.LEFT, padx=5)
tk.Button(frame_topo, text="Calcular", command=calcular_manual, bg="blue", fg="white").pack(side=tk.LEFT, padx=5)

frame_conteudo = tk.Frame(frame_manual)
frame_conteudo.pack()

frame_form = tk.Frame(frame_conteudo)
frame_form.pack(side=tk.LEFT, padx=10)
tk.Label(frame_form, text="Largura (mm):").pack()
entry_largura = tk.Entry(frame_form)
entry_largura.pack()
tk.Label(frame_form, text="Altura (mm):").pack()
entry_altura = tk.Entry(frame_form)
entry_altura.pack()
tk.Label(frame_form, text="Comprimento (m):").pack()
entry_comprimento = tk.Entry(frame_form)
entry_comprimento.pack()
tk.Label(frame_form, text="Espessura (mm):").pack()
entry_espessura = tk.Entry(frame_form)
entry_espessura.pack()
tk.Button(frame_form, text="Adicionar Peça", command=adicionar_peca, bg="orange", fg="black").pack(pady=5)

tree = ttk.Treeview(frame_conteudo, columns=("Largura", "Altura", "Comprimento", "Espessura"), show="headings")
tree.heading("Largura", text="Largura (mm)")
tree.heading("Altura", text="Altura (mm)")
tree.heading("Comprimento", text="Comprimento (m)")
tree.heading("Espessura", text="Espessura (mm)")
tree.pack(side=tk.RIGHT)

label_resultado = tk.Label(frame_manual, text="")
label_resultado.pack()

root.mainloop()

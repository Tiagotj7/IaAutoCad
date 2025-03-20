import xlwings as xw
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import ezdxf

# Tabela de pesos por espessura (kg/m²)
PESO_CHAPA = {
    0.61: 4.88, 0.68: 5.49, 0.76: 6.10, 0.84: 6.71, 0.91: 7.32, 1.06: 8.54,
    1.21: 9.76, 1.37: 10.98, 1.52: 12.21, 1.71: 13.73, 1.90: 15.26, 2.28: 18.81,
    2.66: 21.36, 3.04: 24.41, 3.42: 27.46, 3.80: 30.52, 4.18: 33.57, 4.55: 36.62,
    4.76: 37.35, 6.35: 49.80, 7.94: 62.25, 9.53: 74.70, 12.70: 99.60, 15.88: 124.49,
    19.05: 149.39, 22.23: 174.29, 25.40: 199.19, 26.99: 211.64, 28.58: 224.09,
    30.16: 236.53, 31.75: 249.98, 33.34: 261.43, 34.93: 273.88
}
PRECO_KG = 15.00  # Preço do kg do material

# Variável global para armazenar o estado dos itens salvos
estado_salvo = []

# ---------------------------
# FUNÇÕES DE MANIPULAÇÃO DE PEÇAS MANUAL
# ---------------------------

def salvar_estado():
    """
    Salva o estado atual dos itens da TreeView em uma lista global,
    permitindo reverter (desfazer) alterações posteriormente.
    """
    global estado_salvo
    estado_salvo = []
    for item in tree.get_children():
        estado_salvo.append(tree.item(item, "values"))

def adicionar_peca():
    """
    Adiciona uma peça à tabela (TreeView) usando os valores digitados
    nos campos de entrada (largura, altura, peso, espessura).
    """
    if entry_largura.get() and entry_altura.get() and entry_peso.get() and entry_espessura.get():
        tree.insert("", "end", values=(entry_largura.get(),
                                       entry_altura.get(),
                                       entry_peso.get(),
                                       entry_espessura.get()))
        salvar_estado()  # Salva estado após adicionar a peça
        limpar_campos()

def adicionar_multiplas_pecas():
    """
    Adiciona múltiplas peças (a quantidade é informada em entry_num_pecas)
    com os mesmos valores de largura, altura, peso e espessura.
    """
    try:
        num_pecas = int(entry_num_pecas.get())
        for _ in range(num_pecas):
            if entry_largura.get() and entry_altura.get() and entry_peso.get() and entry_espessura.get():
                tree.insert("", "end", values=(entry_largura.get(),
                                               entry_altura.get(),
                                               entry_peso.get(),
                                               entry_espessura.get()))
        salvar_estado()  # Salva estado após adicionar as peças
        limpar_campos()
    except ValueError:
        messagebox.showerror("Erro", "Número de peças inválido.")

def remover_peca():
    """
    Remove a peça selecionada na tabela (TreeView). Salva o estado antes
    para permitir reverter a remoção.
    """
    item_selecionado = tree.selection()
    if item_selecionado:
        salvar_estado()
        tree.delete(item_selecionado)

def remover_tudo():
    """
    Remove todas as peças da tabela (TreeView). Salva o estado antes
    para permitir reverter a remoção em lote.
    """
    global estado_salvo
    salvar_estado()
    for item in tree.get_children():
        tree.delete(item)

def editar_peca():
    """
    Carrega os valores da peça selecionada para os campos de entrada
    e remove a peça original da tabela. O usuário pode então modificar
    e clicar em "Adicionar Peça" para inserir a versão editada.
    """
    item_selecionado = tree.selection()
    if item_selecionado:
        valores = tree.item(item_selecionado, "values")
        entry_largura.delete(0, tk.END)
        entry_altura.delete(0, tk.END)
        entry_peso.delete(0, tk.END)
        entry_espessura.delete(0, tk.END)

        entry_largura.insert(0, valores[0])
        entry_altura.insert(0, valores[1])
        entry_peso.insert(0, valores[2])
        entry_espessura.insert(0, valores[3])

        salvar_estado()
        tree.delete(item_selecionado)

def reverter():
    """
    Reverte a tabela (TreeView) para o último estado salvo.
    """
    global estado_salvo
    if estado_salvo:
        for item in tree.get_children():
            tree.delete(item)
        for valores in estado_salvo:
            tree.insert("", "end", values=valores)

def limpar_campos():
    """
    Limpa os campos de entrada de dados (largura, altura, peso, espessura, número de peças).
    """
    entry_largura.delete(0, tk.END)
    entry_altura.delete(0, tk.END)
    entry_peso.delete(0, tk.END)
    entry_espessura.delete(0, tk.END)
    entry_num_pecas.delete(0, tk.END)

def calcular_manual():
    """
    Percorre todas as peças listadas na tabela (TreeView), soma o peso total
    e calcula o preço total (Peso x PRECO_KG), exibindo o resultado em label_resultado.
    """
    total_peso = 0
    total_preco = 0
    for item in tree.get_children():
        valores = tree.item(item, "values")
        try:
            # Agora a ordem é: Largura, Altura, Peso e Espessura
            largura, altura, peso, espessura = map(float, valores)
            if espessura not in PESO_CHAPA:
                messagebox.showerror("Erro", f"Espessura {espessura} mm não encontrada na tabela.")
                return
            total_peso += peso
            total_preco += peso * PRECO_KG
        except ValueError:
            messagebox.showerror("Erro", "Erro nos valores inseridos.")
            return

    label_resultado.config(text=f"Peso Total: {total_peso:.2f} kg | Preço Total: R$ {total_preco:.2f}")

# ---------------------------
# FUNÇÕES PARA LEITURA DE DXF E ATUALIZAÇÃO EXCEL
# ---------------------------

def ler_dxf(arquivo_dxf):
    """
    Lê um arquivo DXF usando a biblioteca ezdxf. Para cada LWPOLYLINE
    que tenha ao menos 4 pontos, estima largura e altura; calcula a área
    e, a partir da tabela PESO_CHAPA, determina a espessura mais próxima e o peso.
    Retorna uma lista de listas com [largura, altura, peso, espessura, preço].
    """
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
                    # Cálculo de área em m²
                    area_m2 = (largura / 1000) * (altura / 1000)
                    # Seleciona a espessura mais próxima com base na "altura"
                    espessura = min(PESO_CHAPA.keys(), key=lambda x: abs(x - altura))
                    # Calcula o peso (kg) a partir da área e do peso por m²
                    peso_kg = area_m2 * PESO_CHAPA[espessura]
                    preco_total = round(peso_kg * PRECO_KG, 2)
                    materiais.append([
                        round(largura, 2),      # Largura (mm)
                        round(altura, 2),       # Altura (mm)
                        round(peso_kg, 2),      # Peso (kg)
                        round(espessura, 2),    # Espessura (mm)
                        preco_total            # Preço (R$)
                    ])
        return materiais
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao ler o arquivo DXF.\nErro: {e}")
        return []

def atualizar_planilha(materiais, caminho_excel):
    """
    Abre o arquivo Excel (usando xlwings), limpa a área A2:F100 na planilha 'Orçamento',
    insere o cabeçalho em A1 e os dados da lista 'materiais' a partir de A2.
    Salva e fecha o arquivo.
    """
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
        ws.range("A2:F100").clear_contents()
        ws.range("A1").value = ["Largura (mm)", "Altura (mm)", "Peso (kg)", "Espessura (mm)", "Preço (R$)"]

        if materiais:
            ws.range("A2").value = materiais
            ws.range("C2:C100").number_format = "0.00"
            ws.range("E2:E100").number_format = "R$ #,##0.00"

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

def executar():
    """
    Faz a leitura do caminho de arquivo DXF e Excel informados,
    chama ler_dxf() para extrair os dados e depois atualizar_planilha()
    para preencher o Excel. Por fim, insere os 4 primeiros valores (Largura,
    Altura, Peso, Espessura) na TreeView para visualização rápida.
    """
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

    # Exibe os 4 primeiros campos (Largura, Altura, Peso, Espessura) na TreeView
    for material in materiais:
        tree.insert("", "end", values=material[:4])

# ---------------------------
# FUNÇÕES DE SELEÇÃO DE ARQUIVOS
# ---------------------------

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

# ---------------------------
# CRIAÇÃO DA INTERFACE GRÁFICA (TKINTER)
# ---------------------------

root = tk.Tk()
root.title("Automação de Orçamentos")
root.geometry("800x600")

# PanedWindow para dividir a janela em painel esquerdo e direito
panedwindow = tk.PanedWindow(root, orient=tk.HORIZONTAL)
panedwindow.pack(fill=tk.BOTH, expand=True)

# Painel esquerdo
left_frame = tk.Frame(panedwindow, width=400)
left_frame.pack(fill=tk.Y, padx=10, pady=10)
panedwindow.add(left_frame)

# Seção de seleção de arquivos (DXF, Excel) e botão "Executar"
frame_arquivos = tk.Frame(left_frame)
frame_arquivos.pack()
tk.Label(frame_arquivos, text="Arquivo DXF:").pack()
entry_dxf = tk.Entry(frame_arquivos, width=50)
entry_dxf.pack()
tk.Button(frame_arquivos, text="Procurar", command=selecionar_arquivo_dxf).pack()

tk.Label(frame_arquivos, text="Arquivo Excel:").pack()
entry_excel = tk.Entry(frame_arquivos, width=50)
entry_excel.pack()
tk.Button(frame_arquivos, text="Procurar", command=selecionar_arquivo_excel).pack()

tk.Button(frame_arquivos, text="Executar", command=executar, bg="green", fg="white").pack(pady=10)

# Seção para adicionar peças manualmente
frame_adicionar = tk.Frame(left_frame)
frame_adicionar.pack()

tk.Label(frame_adicionar, text="Número de Peças a Adicionar:").pack()
entry_num_pecas = tk.Entry(frame_adicionar, width=10)
entry_num_pecas.pack()

tk.Button(frame_adicionar, text="Adicionar Múltiplas Peças",
          command=adicionar_multiplas_pecas, bg="green", fg="white").pack(pady=10)

tk.Label(frame_adicionar, text="Largura (mm):").pack()
entry_largura = tk.Entry(frame_adicionar)
entry_largura.pack()

tk.Label(frame_adicionar, text="Altura (mm):").pack()
entry_altura = tk.Entry(frame_adicionar)
entry_altura.pack()

tk.Label(frame_adicionar, text="Peso (kg):").pack()
entry_peso = tk.Entry(frame_adicionar)
entry_peso.pack()

tk.Label(frame_adicionar, text="Espessura (mm):").pack()
entry_espessura = tk.Entry(frame_adicionar)
entry_espessura.pack()

tk.Button(frame_adicionar, text="Adicionar Peça",
          command=adicionar_peca, bg="blue", fg="white").pack(pady=10)

# Painel direito com a TreeView para listar as peças
right_frame = tk.Frame(panedwindow)
right_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
panedwindow.add(right_frame)

tree = ttk.Treeview(right_frame, columns=("Largura", "Altura", "Peso", "Espessura"), show="headings")
tree.heading("Largura", text="Largura (mm)")
tree.heading("Altura", text="Altura (mm)")
tree.heading("Peso", text="Peso (kg)")
tree.heading("Espessura", text="Espessura (mm)")
tree.pack(fill=tk.BOTH, expand=True)

# Botões de ação abaixo da TreeView
button_frame = tk.Frame(right_frame)
button_frame.pack(pady=10)

tk.Button(button_frame, text="Reverter", command=reverter, bg="yellow", fg="black").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Remover Peça", command=remover_peca, bg="red", fg="white").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Remover Tudo", command=remover_tudo, bg="black", fg="white").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Editar Peça", command=editar_peca, bg="orange", fg="black").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Calcular Preço", command=calcular_manual, bg="green", fg="white").pack(side=tk.LEFT, padx=5)

# Label que exibirá o resultado do cálculo de peso/preço
label_resultado = tk.Label(right_frame, text="Peso Total: 0 kg | Preço Total: R$ 0.00")
label_resultado.pack()

root.mainloop()

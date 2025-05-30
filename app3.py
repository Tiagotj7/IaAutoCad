from bokeh.plotting import figure, show
from bokeh.models import ColumnDataSource, HoverTool
from bokeh.layouts import column
from bokeh.io import output_file, curdoc
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import xlwings as xw
import pandas as pd
import os

import ezdxf
from decimal import Decimal, ROUND_HALF_UP, getcontext

# Configura a precisão (suficiente para nossos cálculos)
getcontext().prec = 10

# Tabela de pesos por espessura (kg/m²) – mantida para registro, mas não é utilizada no cálculo atual
PESO_CHAPA = {
    0.61: 4.88, 0.68: 5.49, 0.76: 6.1, 0.84: 6.71, 0.91: 7.32, 1.06: 8.54,
    1.21: 9.76, 1.37: 10.98, 1.52: 12.21, 1.71: 13.73, 1.9: 15.26, 2.28: 18.81,
    2.66: 21.36, 3.04: 24.41, 3.42: 27.46, 3.8: 30.52, 4.18: 33.57, 4.55: 36.62,
    4.76: 37.35, 6.35: 49.8, 7.94: 62.25, 9.53: 74.7, 12.7: 99.6, 15.88: 124.49,
    19.05: 149.39, 22.23: 174.29, 25.4: 199.19, 26.99: 211.64, 28.58: 224.09,
    30.16: 236.53, 31.75: 249.98, 33.34: 261.43, 34.93: 273.88
}

# Fator de preço padrão (será atualizado conforme o material)
# Usaremos Decimal para maior precisão. Inicialmente para inox.
current_preco_kg = Decimal("2.03333")

# Variável global para armazenar o estado dos itens salvos
estado_salvo = []

# Função para criar visualização Bokeh dos dados
def criar_visualizacao():
    # Coleta dados da tabela para visualização
    larguras = []
    alturas = []
    pesos = []
    precos = []
    materiais = []
    
    for item in tree.get_children():
        valores = tree.item(item, "values")
        try:
            largura = float(valores[0])
            altura = float(valores[1])
            peso = float(valores[2])
            material = valores[4]
            
            # Calculando o preço baseado no material
            if material == "Inox":
                preco_fator = float(Decimal("2.03333"))
            elif material == "Aço":
                preco_fator = float(Decimal("2.878205128205128205"))
            else:  # Alumínio
                preco_fator = float(Decimal("4.5"))
                
            preco = peso * preco_fator
            
            larguras.append(largura)
            alturas.append(altura)
            pesos.append(peso)
            precos.append(preco)
            materiais.append(material)
        except Exception as e:
            print(f"Erro ao processar item: {e}")
    
    # Criar fonte de dados para Bokeh
    source = ColumnDataSource(data={
        'largura': larguras,
        'altura': alturas,
        'peso': pesos,
        'preco': precos,
        'material': materiais,
        'area': [l*a/1000000 for l, a in zip(larguras, alturas)]  # área em m²
    })
    
    # Criar gráfico de dispersão para área vs preço
    p1 = figure(title="Área vs Preço", x_axis_label="Área (m²)", y_axis_label="Preço (R$)",
               width=600, height=400, toolbar_location="above")
    
    # Diferentes cores para diferentes materiais
    colors = {"Inox": "blue", "Aço": "red", "Alumínio": "green"}
    
    # Adicionar círculos coloridos por material
    for material, color in colors.items():
        source_filtered = ColumnDataSource(data={
            'area': [a for a, m in zip(source.data['area'], source.data['material']) if m == material],
            'preco': [p for p, m in zip(source.data['preco'], source.data['material']) if m == material],
            'largura': [l for l, m in zip(source.data['largura'], source.data['material']) if m == material],
            'altura': [a for a, m in zip(source.data['altura'], source.data['material']) if m == material],
            'peso': [p for p, m in zip(source.data['peso'], source.data['material']) if m == material],
        })
        
        if len(source_filtered.data['area']) > 0:  # Verifica se há dados para este material
            p1.circle('area', 'preco', size=10, color=color, legend_label=material, source=source_filtered)
    
    # Adicionar ferramenta de hover para mostrar detalhes ao passar o mouse
    hover = HoverTool()
    hover.tooltips = [
        ("Largura", "@largura mm"),
        ("Altura", "@altura mm"),
        ("Área", "@area m²"),
        ("Peso", "@peso kg"),
        ("Preço", "R$ @preco{0,0.00}")
    ]
    p1.add_tools(hover)
    
    # Criar gráfico de barras para peso por material
    peso_por_material = {}
    for material in set(materiais):
        peso_por_material[material] = sum(p for p, m in zip(pesos, materiais) if m == material)
    
    p2 = figure(title="Peso Total por Material", x_range=list(peso_por_material.keys()), 
                height=400, width=600, toolbar_location="above")
    
    p2.vbar(x=list(peso_por_material.keys()), top=list(peso_por_material.values()), 
            width=0.5, color=[colors.get(m, "gray") for m in peso_por_material.keys()])
    
    p2.xgrid.grid_line_color = None
    p2.y_range.start = 0
    p2.yaxis.axis_label = "Peso Total (kg)"
    
    # Salvar os gráficos em um arquivo HTML
    output_file("visualizacao_chapas.html")
    
    # Retornar os gráficos para exibição
    return column(p1, p2)

# Função para exibir a visualização Bokeh
def mostrar_visualizacao():
    if len(tree.get_children()) == 0:
        messagebox.showwarning("Aviso", "Não há dados para visualizar.")
        return
    
    # Criar e exibir a visualização
    graficos = criar_visualizacao()
    show(graficos)  # Isso abrirá a visualização no navegador

# Funções específicas para cada material
def func_inox():
    global current_preco_kg
    # Para inox, fator = 13.42/6.6 ≈ 2.03333
    current_preco_kg = Decimal("2.03333")
    entry_largura.delete(0, tk.END)
    entry_largura.insert(0, "1000")
    entry_altura.delete(0, tk.END)
    entry_altura.insert(0, "2000")
    entry_peso.delete(0, tk.END)
    entry_peso.insert(0, "6.6")
    entry_espessura.delete(0, tk.END)
    entry_espessura.insert(0, "0.84")
    calcular_manual()

def func_aco():
    global current_preco_kg
    # Para aço, fator = 22.46/7.8 ≈ 2.878205128205128205, usamos mais dígitos para precisão exata.
    current_preco_kg = Decimal("2.878205128205128205")
    entry_largura.delete(0, tk.END)
    entry_largura.insert(0, "1200")
    entry_altura.delete(0, tk.END)
    entry_altura.insert(0, "2400")
    entry_peso.delete(0, tk.END)
    entry_peso.insert(0, "7.8")
    entry_espessura.delete(0, tk.END)
    entry_espessura.insert(0, "1.00")
    calcular_manual()

def func_aluminio():
    global current_preco_kg
    # Para alumínio, fator = 18.45/4.1 ≈ 4.5
    current_preco_kg = Decimal("4.5")
    entry_largura.delete(0, tk.END)
    entry_largura.insert(0, "1500")
    entry_altura.delete(0, tk.END)
    entry_altura.insert(0, "3000")
    entry_peso.delete(0, tk.END)
    entry_peso.insert(0, "4.1")
    entry_espessura.delete(0, tk.END)
    entry_espessura.insert(0, "1.50")
    calcular_manual()

def material_selected(event):
    material = combo_material.get()
    if material == "Inox":
        func_inox()
    elif material == "Aço":
        func_aco()
    elif material == "Alumínio":
        func_aluminio()

# Função para adicionar peça na tabela manual
def adicionar_peca():
    if entry_largura.get() and entry_altura.get() and entry_peso.get() and entry_espessura.get() and combo_material.get():
        tree.insert("", "end", values=(entry_largura.get(), entry_altura.get(), entry_peso.get(), entry_espessura.get(), combo_material.get()))
        salvar_estado()  # Salvar estado após adicionar a peça
        limpar_campos()

# Função para adicionar múltiplas peças
def adicionar_multiplas_pecas():
    num_pecas = entry_num_pecas.get()
    try:
        num_pecas = int(num_pecas)
        for _ in range(num_pecas):
            if entry_largura.get() and entry_altura.get() and entry_peso.get() and entry_espessura.get() and combo_material.get():
                tree.insert("", "end", values=(entry_largura.get(), entry_altura.get(), entry_peso.get(), entry_espessura.get(), combo_material.get()))
        salvar_estado()  # Salvar estado após adicionar as peças
        limpar_campos()
    except ValueError:
        messagebox.showerror("Erro", "Número de peças inválido.")

# Função para remover peça da tabela manual
def remover_peca():
    item_selecionado = tree.selection()
    if item_selecionado:
        salvar_estado()  # Salvar estado antes de remover a peça
        tree.delete(item_selecionado)

# Função para salvar o estado dos itens
def salvar_estado():
    global estado_salvo
    estado_salvo = []
    for item in tree.get_children():
        estado_salvo.append(tree.item(item, "values"))

# Função para remover tudo da tabela manual
def remover_tudo():
    global estado_salvo
    salvar_estado()  # Salvar o estado antes de remover todos os itens
    for item in tree.get_children():
        tree.delete(item)

# Função para editar peça na tabela manual
def editar_peca():
    item_selecionado = tree.selection()
    if item_selecionado:
        valores = tree.item(item_selecionado, "values")
        entry_largura.delete(0, tk.END)
        entry_altura.delete(0, tk.END)
        entry_peso.delete(0, tk.END)
        entry_espessura.delete(0, tk.END)
        combo_material.set('')
        entry_largura.insert(0, valores[0])
        entry_altura.insert(0, valores[1])
        entry_peso.insert(0, valores[2])
        entry_espessura.insert(0, valores[3])
        combo_material.set(valores[4])
        salvar_estado()  # Salvar o estado antes de editar a peça
        tree.delete(item_selecionado)

# Função para calcular preço manualmente usando Decimal
def calcular_manual():
    total_preco = Decimal("0.00")
    for item in tree.get_children():
        valores = tree.item(item, "values")
        try:
            largura, altura, peso, espessura = map(Decimal, valores[:4])
            preco_total = (peso * current_preco_kg).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            total_preco += preco_total
        except Exception:
            messagebox.showerror("Erro", "Erro nos valores inseridos.")
            return
    label_resultado.config(text=f"Preço Total: R$ {total_preco}")

# Função para limpar campos
def limpar_campos():
    entry_largura.delete(0, tk.END)
    entry_altura.delete(0, tk.END)
    entry_peso.delete(0, tk.END)
    entry_espessura.delete(0, tk.END)
    entry_num_pecas.delete(0, tk.END)
    combo_material.set('')

# Função para remover tudo da tabela manual (mesma que remover_tudo)
def remover():
    global estado_salvo
    salvar_estado()
    for item in tree.get_children():
        tree.delete(item)

# Função para reverter tudo (trazer de volta todos os itens salvos)
def reverter():
    global estado_salvo
    if estado_salvo:
        for item in tree.get_children():
            tree.delete(item)
        for valores in estado_salvo:
            tree.insert("", "end", values=valores)

# Função para ler arquivo DXF
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
                    # Aqui, 'peso' é lido diretamente (supostamente em kg) a partir da elevação do DXF
                    peso = entidade.dxf.elevation
                    espessura = min(PESO_CHAPA.keys(), key=lambda x: abs(x - altura))
                    preco_total = (Decimal(str(peso)) * current_preco_kg).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    # Para DXF, atribuímos "Inox" como material padrão
                    materiais.append([round(largura, 2), round(altura, 2), round(peso, 2), round(espessura, 2), "Inox", float(preco_total)])
        return materiais
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao ler o arquivo DXF.\nErro: {e}")
        return []

# Função para atualizar a planilha Excel (incluindo o campo Material)
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
        # Cabeçalho atualizado para incluir Material
        ws.range("A1").value = ["Largura (cm)", "Altura (cm)", "Peso (kg)", "Espessura (mm)", "Material", "Preço (R$)"]
        if materiais:
            ws.range("A2").value = materiais
            ws.range("F2:F100").number_format = "R$ #,##0.00"
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

# Função para executar a leitura DXF e atualização do Excel
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
    for material in materiais:
        # Inserindo as cinco colunas: Largura, Altura, Peso, Espessura e Material
        tree.insert("", "end", values=material[:5])

# Função para selecionar arquivo DXF
def selecionar_arquivo_dxf():
    arquivo = filedialog.askopenfilename(filetypes=[("Arquivos DXF", "*.dxf")])
    if arquivo:
        entry_dxf.delete(0, tk.END)
        entry_dxf.insert(0, arquivo)

# Função para selecionar arquivo Excel
def selecionar_arquivo_excel():
    arquivo = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx")])
    if arquivo:
        entry_excel.delete(0, tk.END)
        entry_excel.insert(0, arquivo)

# Interface gráfica
root = tk.Tk()
root.title("Automação de Orçamentos")
root.geometry("800x600")

# Divisão da janela em dois painéis
panedwindow = tk.PanedWindow(root, orient=tk.HORIZONTAL)
panedwindow.pack(fill=tk.BOTH, expand=True)

# Painel esquerdo para adicionar peça personalizada
left_frame = tk.Frame(panedwindow, width=400)
left_frame.pack(fill=tk.Y, padx=10, pady=10)
panedwindow.add(left_frame)

# Frame para seleção de arquivos e dados manual
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

# Frame para adicionar peças manualmente à esquerda
frame_adicionar = tk.Frame(left_frame)
frame_adicionar.pack()

tk.Label(frame_adicionar, text="Número de Peças a Adicionar:").pack()
entry_num_pecas = tk.Entry(frame_adicionar, width=10)
entry_num_pecas.pack()
tk.Button(frame_adicionar, text="Adicionar Múltiplas Peças", command=adicionar_multiplas_pecas, bg="green", fg="white").pack(pady=10)

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

tk.Label(frame_adicionar, text="Material:").pack()
combo_material = ttk.Combobox(frame_adicionar, values=["Inox", "Aço", "Alumínio"])
combo_material.set("Inox")
combo_material.pack()
combo_material.bind("<<ComboboxSelected>>", material_selected)

tk.Button(frame_adicionar, text="Adicionar Peça", command=adicionar_peca, bg="blue", fg="white").pack(pady=10)

# Painel direito para tabela de peças
right_frame = tk.Frame(panedwindow)
right_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
panedwindow.add(right_frame)

tree = ttk.Treeview(right_frame, columns=("Largura", "Altura", "Peso", "Espessura", "Material"), show="headings")
tree.heading("Largura", text="Largura (mm)")
tree.heading("Altura", text="Altura (mm)")
tree.heading("Peso", text="Peso (kg)")
tree.heading("Espessura", text="Espessura (mm)")
tree.heading("Material", text="Material")
tree.pack(fill=tk.BOTH, expand=True)

# Botões de ação
button_frame = tk.Frame(right_frame)
button_frame.pack(pady=10)
tk.Button(button_frame, text="Reverter", command=reverter, bg="yellow", fg="black").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Remover Peça", command=remover_peca, bg="red", fg="white").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Remover Tudo", command=remover_tudo, bg="black", fg="white").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Editar Peça", command=editar_peca, bg="orange", fg="black").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Calcular Preço", command=calcular_manual, bg="green", fg="white").pack(side=tk.LEFT, padx=5)
tk.Button(button_frame, text="Visualizar Gráficos", command=mostrar_visualizacao, bg="purple", fg="white").pack(side=tk.LEFT, padx=5)

label_resultado = tk.Label(right_frame, text="Preço Total: R$ 0.00")
label_resultado.pack()

root.mainloop()
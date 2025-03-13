# Automação de Orçamentos com DXF e Excel

🚀 Automatize a criação de orçamentos a partir de arquivos DXF! Este projeto utiliza Python, AutoCAD (DXF) e Excel para extrair dados de dutos, calcular materiais e preencher automaticamente uma planilha de orçamento.

## 📌 Funcionalidades
- ✅ **Leitura de Arquivos DXF**: Extração automática das dimensões de materiais a partir de desenhos no formato DXF.
- ✅ **Cálculo de Materiais**: Determina a área, peso e custo com base em valores pré-definidos para espessura do material.
- ✅ **Atualização Automática do Excel**: Insere os dados calculados diretamente na planilha de orçamento do Excel.
- ✅ **Interface Amigável**: Entrada via terminal ou interface gráfica para fácil interação com os arquivos DXF e Excel.
- ✅ **Tabela Interativa de Peças**: Gerencie as peças inseridas manualmente, com possibilidade de adicionar, editar, remover e calcular preços.

## 🛠 Tecnologias Utilizadas
- **Python**
- **ezdxf** → Para leitura de arquivos DXF
- **xlwings** → Para manipulação do Excel
- **pandas** → Para tratamento de dados
- **tkinter** → Para a interface gráfica (opcional)
- **ttk (Treeview)** → Para exibição interativa de tabelas de peças

## 🚀 Instalação e Uso

### 1️⃣ Instale as Dependências

Execute o seguinte comando para instalar as bibliotecas necessárias:

```bash
pip install ezdxf xlwings pandas openpyxl

Para Linux (caso necessário), instale o tkinter com:
sudo apt install python3-tk

2️⃣ Execute o Script
Com interface de terminal:
python app.py --dxf "caminho/do/arquivo.dxf" --excel "caminho/da/planilha.xlsx"

Ou utilize a interface gráfica para selecionar os arquivos DXF e Excel diretamente.
📊 Exemplo de Saída no Excel
Largura (cm)
Altura (cm)
Comprimento (m)
Espessura (mm)
Peso (kg)
Preço (R$)
40
30
4
1.5
37.68
565.20
50
40
5
2.0
78.50
1177.50

💡 Possíveis Melhorias
📌 Suporte para mais tipos de materiais (ex.: tubos de cobre).
📌 Integração com banco de dados para histórico de orçamentos.
📌 Interface gráfica mais avançada para facilitar o uso.
📜 Licença
Este projeto está sob a licença MIT. Sinta-se livre para contribuir e aprimorar!

👷 Desenvolvido para otimizar orçamentos no setor de construção e refrigeração! 🚀

Este formato segue a estrutura típica de um arquivo `README.md`, com seções claras para funcionalidades, instalação, exemplo de saída, melhorias e licenciamento.



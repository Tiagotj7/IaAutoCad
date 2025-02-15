# **📐 IA para Automação de Orçamentos no AutoCAD**  

🚀 **Automatize a geração de orçamentos de obras a partir de arquivos DWG!** Este projeto utiliza **Python**, **AutoCAD (DWG)** e **Excel** para extrair medidas de dutos, calcular materiais e preencher automaticamente uma planilha de orçamento.  

## **📌 Funcionalidades**  
✅ **Leitura de Arquivos DWG**: Extração automática das dimensões dos dutos a partir de desenhos do AutoCAD.  
✅ **Cálculo de Materiais**: Determina a área, peso e custo com base em valores pré-definidos.  
✅ **Atualização Automática do Excel**: Insere os dados calculados diretamente na planilha de orçamento.  
✅ **Interface Amigável**: Possibilidade de entrada via terminal ou interface gráfica para seleção dos arquivos.  

## **🛠 Tecnologias Utilizadas**  
- **Python**  
- **ezdxf** → Para leitura de arquivos DWG  
- **xlwings** → Para manipulação do Excel  
- **pandas** → Para tratamento de dados  
- **tkinter** → Para a interface gráfica (opcional)  

## **🚀 Instalação e Uso**  

### **1️⃣ Instale as Dependências**  
```bash
pip install ezdxf xlwings pandas openpyxl
```
Para Linux (caso necessário):  
```bash
sudo apt install python3-tk
```

### **2️⃣ Execute o Script**
Com interface de terminal:  
```bash
python app.py --dwg "caminho/do/arquivo.dwg" --excel "caminho/da/planilha.xlsx"
```
Ou selecione os arquivos via interface gráfica.  

## **📊 Exemplo de Saída no Excel**  

| Largura (cm) | Altura (cm) | Comprimento (m) | Área (m²) | Peso (kg) | Preço (R$) |
|-------------|-------------|----------------|-----------|-----------|-----------|
| 40         | 30         | 4              | 4.8       | 37.68     | 565.20    |
| 50         | 40         | 5              | 10.0      | 78.50     | 1177.50   |

## **💡 Possíveis Melhorias**  
📌 Suporte para mais tipos de materiais (ex.: tubos de cobre).  
📌 Integração com banco de dados para histórico de orçamentos.  
📌 Interface gráfica mais avançada para facilitar o uso.  

## **📜 Licença**  
Este projeto está sob a licença **MIT**. Sinta-se livre para contribuir e aprimorar!  

👷 **Desenvolvido para otimizar orçamentos no setor de refrigeração e construção!** 🚀
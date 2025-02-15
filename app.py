import ezdxf
import xlwings as xw
import pandas as pd
import os

# Definições de cálculo (Exemplo: Peso da chapa por metro quadrado)
PESO_CHAPA_KG_M2 = 7.85  # Exemplo de densidade da chapa
PRECO_KG = 15.00  # Preço por kg da chapa

def ler_dwg(arquivo_dwg):
    """
    Lê o arquivo DWG e extrai as dimensões dos dutos.
    Retorna uma lista de dicionários com largura, altura, comprimento e peso estimado.
    """
    doc = ezdxf.readfile(arquivo_dwg)
    msp = doc.modelspace()

    materiais = []

    for entidade in msp.query("LWPOLYLINE"):  # Captura apenas polilinhas (pode ajustar conforme o DWG)
        if len(entidade.get_points()) >= 4:  # Garante que é um retângulo
            pontos = entidade.get_points()
            largura = abs(pontos[1][0] - pontos[0][0])  # Diferença X
            altura = abs(pontos[2][1] - pontos[1][1])  # Diferença Y
            comprimento = entidade.dxf.elevation  # Supondo que a elevação seja o comprimento

            # Cálculo da área total em metros quadrados
            area_m2 = (largura / 1000) * (altura / 1000) * comprimento  # Convertendo mm para metros

            # Peso estimado
            peso_kg = area_m2 * PESO_CHAPA_KG_M2

            # Preço estimado
            preco_total = peso_kg * PRECO_KG

            materiais.append({
                "Largura (cm)": largura,
                "Altura (cm)": altura,
                "Comprimento (m)": comprimento,
                "Área (m²)": round(area_m2, 2),
                "Peso (kg)": round(peso_kg, 2),
                "Preço (R$)": round(preco_total, 2)
            })

    return materiais

def atualizar_planilha(materiais, caminho_excel):
    """
    Atualiza a planilha do Excel com os materiais extraídos do DWG.
    """
    if not os.path.exists(caminho_excel):
        print("Erro: Arquivo Excel não encontrado.")
        return
    
    app = xw.App(visible=False)  # Abre Excel em segundo plano
    wb = xw.Book(caminho_excel)
    ws = wb.sheets["Orçamento"]  # Nome da aba a ser preenchida

    # Limpar dados antigos antes de inserir os novos
    ws.range("A2:F100").clear_contents()

    # Escrever cabeçalhos
    ws.range("A1").value = ["Largura (cm)", "Altura (cm)", "Comprimento (m)", "Área (m²)", "Peso (kg)", "Preço (R$)"]

    # Preencher os dados
    ws.range("A2").value = [list(mat.values()) for mat in materiais]

    wb.save()
    wb.close()
    app.quit()

    print("Planilha atualizada com sucesso!")

def main():
    """
    Executa o fluxo completo do programa.
    """
    arquivo_dwg = input("Digite o caminho do arquivo DWG: ")
    caminho_excel = input("Digite o caminho do arquivo Excel: ")

    if not os.path.exists(arquivo_dwg):
        print("Erro: Arquivo DWG não encontrado!")
        return

    materiais = ler_dwg(arquivo_dwg)

    if not materiais:
        print("Nenhum material encontrado no DWG.")
        return

    atualizar_planilha(materiais, caminho_excel)

if __name__ == "__main__":
    main()

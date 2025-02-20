from flask import Flask, request, jsonify, send_file
import xlwings as xw
import os
import ezdxf

app = Flask(__name__)

PESO_CHAPA_KG_M2 = 7.85
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
        return str(e)

def atualizar_planilha(materiais, caminho_excel):
    if not os.path.exists(caminho_excel):
        return "Arquivo Excel não encontrado."

    try:
        app = xw.App(visible=False)
        wb = xw.Book(caminho_excel)
        if "Orçamento" not in [sheet.name for sheet in wb.sheets]:
            wb.close()
            app.quit()
            return 'A planilha "Orçamento" não foi encontrada no arquivo Excel.'

        ws = wb.sheets["Orçamento"]
        ws.range("A2:F100").clear_contents()
        ws.range("A1").value = ["Largura (cm)", "Altura (cm)", "Comprimento (m)", "Área (m²)", "Peso (kg)", "Preço (R$)"]

        if materiais:
            dados = [list(mat.values()) for mat in materiais]
            ws.range("A2").value = dados

        wb.save()
        wb.close()
        app.quit()
        return "Planilha atualizada com sucesso!"
    except Exception as e:
        return str(e)

@app.route('/processar', methods=['POST'])
def processar_arquivo():
    if 'arquivo_dxf' not in request.files or 'arquivo_excel' not in request.files:
        return jsonify({"erro": "Arquivos não encontrados"}), 400

    arquivo_dxf = request.files['arquivo_dxf']
    arquivo_excel = request.files['arquivo_excel']

    caminho_dxf = "temp.dxf"
    caminho_excel = "planilha.xlsx"

    arquivo_dxf.save(caminho_dxf)
    arquivo_excel.save(caminho_excel)

    materiais = ler_dxf(caminho_dxf)
    if not materiais:
        return jsonify({"erro": "Nenhum material encontrado no DXF"}), 400

    resultado = atualizar_planilha(materiais, caminho_excel)
    return jsonify({"mensagem": resultado})

if __name__ == '__main__':
    app.run(debug=True)

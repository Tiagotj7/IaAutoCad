import os
import ezdxf
import xlwings as xw
from flask import Flask, request, jsonify

PESO_CHAPA_KG_M2 = 7.85
PRECO_KG = 15.00

app = Flask(__name__)

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

@app.route('/api/processar', methods=['POST'])
def processar_arquivo():
    if 'arquivo_dxf' not in request.files or 'arquivo_excel' not in request.files:
        return jsonify({"erro": "Arquivos não encontrados"}), 400

    arquivo_dxf = request.files['arquivo_dxf']
    arquivo_excel = request.files['arquivo_excel']

    caminho_dxf = "/tmp/temp.dxf"
    caminho_excel = "/tmp/planilha.xlsx"

    arquivo_dxf.save(caminho_dxf)
    arquivo_excel.save(caminho_excel)

    materiais = ler_dxf(caminho_dxf)
    if not materiais:
        return jsonify({"erro": "Nenhum material encontrado no DXF"}), 400

    return jsonify({"mensagem": "Processamento concluído!", "materiais": materiais})

if __name__ == '__main__':
    app.run()

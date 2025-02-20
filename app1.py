from flask import Flask, render_template, request, jsonify
import os
import ezdxf  # Para manipular arquivos DXF
import xlwings as xw

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')  # Cria um arquivo HTML para a interface

@app.route('/processar', methods=['POST'])
def processar_arquivo():
    arquivo_dxf = request.files['arquivo_dxf']
    caminho_excel = 'orcamento.xlsx'

    if not arquivo_dxf:
        return jsonify({'erro': 'Nenhum arquivo DXF enviado'})

    # Salva temporariamente o arquivo
    arquivo_dxf.save('temp.dxf')

    # Processa o DXF
    materiais = ler_dxf('temp.dxf')
    if not materiais:
        return jsonify({'erro': 'Nenhum material encontrado'})

    # Atualiza a planilha
    atualizar_planilha(materiais, caminho_excel)

    return jsonify({'mensagem': 'Processamento concluído com sucesso!'})

if __name__ == '__main__':
    app.run(debug=True)

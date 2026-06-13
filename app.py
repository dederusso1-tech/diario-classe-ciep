import os
import pandas as pd
from io import BytesIO
from flask import Flask, render_template_string, send_file, request, redirect, url_for

app = Flask(__name__)

# Dados oficiais fixos dos alunos do CIEP 205
def obtener_dados_alunos():
    return [
        {"nome": "ALLAN ARAÚJO MARQUES DA CONCEIÇÃO", "faltas": 0, "media": 0.0, "d10": "P", "d11": "P"},
        {"nome": "ALLYCYA SANNY GONÇALVES MENESES", "faltas": 0, "media": 0.0, "d10": "P", "d11": "P"},
        {"nome": "ANA BEATRIZ DIAS SILVA", "faltas": 0, "media": 0.0, "d10": "P", "d11": "P"},
        {"nome": "ANA BEATRIZ SANTOS DA PENHA", "faltas": 0, "media": 0.0, "d10": "-", "d11": "-"},
        {"nome": "ANA CLARA SANTOS", "faltas": 0, "media": 0.0, "d10": "P", "d11": "P"}
    ]

# Interface visual em HTML/Bootstrap (Unificada dentro do próprio código)
HTML_DIARIO = '''
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Diário Eletrônico Oficial</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        body { background-color: #f8f9fa; }
        .navbar-custom { background-color: #1a365d; color: white; }
        .card-header-custom { background-color: #fff; border-left: 5px solid #ecc94b; }
        .table-thead { background-color: #f7fafc; }
        .text-p { color: #2f855a; font-weight: bold; }
    </style>
</head>
<body>

<nav class="navbar navbar-custom p-3 mb-4">
    <div class="container-fluid d-flex justify-content-between">
        <span class="navbar-brand mb-0 h1 text-white">🍎 Diário Eletrônico de Classe</span>
        <button class="btn btn-outline-light btn-sm">Sair do Sistema</button>
    </div>
</nav>

<div class="container bg-white p-4 rounded shadow-sm mb-4">
    <div class="alert alert-info alert-dismissible fade show p-2" role="alert">
        Diário de classe disponível!
    </div>

    <p class="text-muted small">Início / CIEP 205 FREI AGOSTINHO FÍNCIAS</p>

    <div class="card card-header-custom p-3 mb-4 shadow-sm">
        <h6 class="fw-bold mb-3">📥 Carga de Dados do Sistema Central (SEEDUC-RJ)</h6>
        <div class="row g-2">
            <div class="col-md-9">
                <input class="form-control" type="file" id="formFile">
            </div>
            <div class="col-md-3">
                <button class="btn btn-warning w-100 fw-bold">Processar Matrículas</button>
            </div>
        </div>
    </div>

    <div class="d-flex justify-content-between align-items-center mb-3">
        <h5 class="fw-bold m-0">📊 PLANILHA CONSOLIDADA (FREQUÊNCIA E RENDIMENTO)</h5>
        <a href="/exportar-excel" class="btn btn-success fw-bold shadow-sm px-4 py-2">
            📥 Baixar Planilha Excel
        </a>
    </div>

    <div class="table-responsive">
        <table class="table table-bordered align-middle">
            <thead class="table-thead">
                <tr>
                    <th>Turma / Nome do Aluno</th>
                    <th class="text-center" style="width: 100px;">Faltas</th>
                    <th class="text-center" style="width: 100px;">Média</th>
                    <th class="text-center" style="width: 80px;">10/06</th>
                    <th class="text-center" style="width: 80px;">11/06</th>
                </tr>
            </thead>
            <tbody>
                {% for aluno in alunos %}
                <tr>
                    <td>
                        <span class="badge bg-secondary mb-1" style="font-size: 10px;">turma</span><br>
                        <strong>{{ aluno.nome }}</strong>
                    </td>
                    <td class="text-center">{{ aluno.faltas }}</td>
                    <td class="text-center text-danger fw-bold">{{ aluno.media }}</td>
                    <td class="text-center text-p">{{ aluno.d10 }}</td>
                    <td class="text-center text-p">{{ aluno.d11 }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>

</body>
</html>
'''

# Rota raiz (redireciona automaticamente para a escola)
@app.route('/')
def index():
    return redirect(url_for('escola', id_escola=1))

# Rota exata visualizada no navegador do Render (/escola/1)
@app.route('/escola/<int:id_escola>')
def escola(id_escola):
    alunos = obtener_dados_alunos()
    return render_template_string(HTML_DIARIO, alunos=alunos)

# Rota que gera e faz o download do arquivo Excel real
@app.route('/exportar-excel')
def exportar_excel():
    alunos = obtener_dados_alunos()
    dados_planilha = []
    
    for aluno in alunos:
        dados_planilha.append({
            "Nome do Aluno": aluno["nome"],
            "Faltas": aluno["faltas"],
            "Média": aluno["media"],
            "10/06": aluno["d10"],
            "11/06": aluno["d11"]
        })
    
    # Criando o arquivo Excel em memória usando o Pandas
    df = pd.DataFrame(dados_planilha)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='CIEP_205')
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Consolidado_Bimestre_CIEP_205.xlsx'
    )

if __name__ == '__main__':
    # Configuração necessária para rodar corretamente em servidores como o Render
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

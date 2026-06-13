import os
import pandas as pd
from io import BytesIO
from flask import Flask, render_template_string, send_file, request, redirect, url_for

app = Flask(__name__)

# Dados oficiais dos alunos da Turma 1017 - CIEP 321
def obtener_dados_alunos():
    return [
        {"nome": "DAVI DOS SANTOS DE FREITAS LOPES PIMENTA", "faltas": 0, "media": 0.0},
        {"nome": "EDUARDA FERREIRA AZEREDO", "faltas": 0, "media": 0.0},
        {"nome": "ELIAS VANDO DO NASCIMENTO SILVA", "faltas": 0, "media": 0.0},
        {"nome": "ISABELLE REJANE SOARES RÊGO", "faltas": 0, "media": 0.0},
        {"nome": "ISABELY LUARA RODRIGUES", "faltas": 0, "media": 0.0},
        {"nome": "ISADORA LETÍCIA DA SILVA MARTINS", "faltas": 0, "media": 0.0},
        {"nome": "KAIO DE ABREU CAMARGO", "faltas": 0, "media": 0.0},
        {"nome": "KAUÃ JIHAD FACCINI SANT'ANA DA SILVA", "faltas": 0, "media": 0.0}
    ]

HTML_DIARIO = '''
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Diário Eletrônico - CIEP 321</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        body { background-color: #f8f9fa; }
        .navbar-custom { background-color: #1a365d; color: white; }
        .table-thead { background-color: #f7fafc; }
    </style>
</head>
<body>

<nav class="navbar navbar-custom p-3 mb-4">
    <div class="container-fluid">
        <span class="navbar-brand mb-0 h1 text-white">🍎 CIEP 321 DR. ULYSSES GUIMARÃES</span>
    </div>
</nav>

<div class="container bg-white p-4 rounded shadow-sm mb-4">
    <p class="text-muted small">Professor: Elisa Carvalho / <strong>Turma: 1017</strong> / Disciplina: Redação</p>

    <div class="d-flex justify-content-between align-items-center mb-3">
        <h5 class="fw-bold m-0">📊 PLANILHA CONSOLIDADA (TURMA 1017)</h5>
        <a href="/exportar-excel" class="btn btn-success fw-bold shadow-sm px-4 py-2">
            📥 Baixar Planilha Excel
        </a>
    </div>

    <div class="table-responsive">
        <table class="table table-bordered align-middle">
            <thead class="table-thead">
                <tr>
                    <th>Nome Completo</th>
                    <th class="text-center" style="width: 100px;">Faltas</th>
                    <th class="text-center" style="width: 100px;">Média</th>
                </tr>
            </thead>
            <tbody>
                {% for aluno in alunos %}
                <tr>
                    <td><strong>{{ aluno.nome }}</strong></td>
                    <td class="text-center">{{ aluno.faltas }}</td>
                    <td class="text-center fw-bold">{{ aluno.media }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</div>

</body>
</html>
'''

@app.route('/')
def index():
    alunos = obtener_dados_alunos()
    return render_template_string(HTML_DIARIO, alunos=alunos)

@app.route('/exportar-excel')
def exportar_excel():
    alunos = obtener_dados_alunos()
    dados_planilha = []
    for aluno in alunos:
        dados_planilha.append({
            "Nome Completo": aluno["nome"],
            "Faltas": aluno["faltas"],
            "Média": aluno["media"]
        })
    
    df = pd.DataFrame(dados_planilha)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='TURMA_1017')
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Consolidado_Turma_1017_CIEP_321.xlsx'
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

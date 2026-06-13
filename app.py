import os
import pandas as pd
from io import BytesIO
from flask import Flask, render_template_string, send_file, request, redirect, url_for

app = Flask(__name__)

# Lista Oficial de Alunos da Turma 1017 (Sem precisar digitar!)
ALUNOS_PADRAO = [
    "DAVI DOS SANTOS DE FREITAS LOPES PIMENTA",
    "EDUARDA FERREIRA AZEREDO",
    "ELIAS VANDO DO NASCIMENTO SILVA",
    "ISABELLE REJANE SOARES RÊGO",
    "ISABELY LUARA RODRIGUES",
    "ISADORA LETÍCIA DA SILVA MARTINS",
    "KAIO DE ABREU CAMARGO",
    "KAUÃ JIHAD FACCINI SANT'ANA DA SILVA"
]

HTML_SISTEMA = '''
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Sistema de Diários - CIEP</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        body { background-color: #f4f6f9; }
        .navbar-custom { background-color: #1a365d; color: white; }
        .card-custom { border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
    </style>
</head>
<body>

<nav class="navbar navbar-custom p-3 mb-4">
    <div class="container d-flex justify-content-between">
        <span class="navbar-brand mb-0 h1 text-white">🍎 Portal do Docente - CIEP</span>
        <span class="text-white-50">Ambiente de Produção</span>
    </div>
</nav>

<div class="container">
    <div class="card card-custom p-4 bg-white mb-4">
        <div class="d-flex justify-content-between align-items-center mb-4">
            <div>
                <h3 class="fw-bold m-0">📊 Diário Eletrônico - Turma 1017</h3>
                <span class="text-muted">Disciplina: Redação | Escola: CIEP 321</span>
            </div>
            <a href="/exportar-excel" class="btn btn-success fw-bold px-4 py-2 shadow-sm">
                📥 Baixar Planilha Excel (Diário)
            </a>
        </div>

        <h5 class="fw-bold mb-3 text-secondary">📋 Lista de Alunos Carregada Automatizada</h5>
        <div class="table-responsive">
            <table class="table table-striped table-bordered align-middle">
                <thead class="table-dark">
                    <tr>
                        <th>Nome Completo do Aluno</th>
                        <th class="text-center" style="width: 150px;">Status</th>
                    </tr>
                </thead>
                <tbody>
                    {% for aluno in alunos %}
                    <tr>
                        <td><strong>{{ aluno }}</strong></td>
                        <td class="text-center"><span class="badge bg-success">Matriculado</span></td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
</div>

</body>
</html>
'''

@app.route('/')
def index():
    # Carrega a página direto com a lista de alunos e o botão de baixar
    return render_template_string(HTML_SISTEMA, alunos=ALUNOS_PADRAO)

@app.route('/exportar-excel')
def exportar_excel():
    # Gera o arquivo Excel instantaneamente com as colunas certas
    dados_planilha = []
    for aluno in ALUNOS_PADRAO:
        dados_planilha.append({
            "Nome Completo": aluno,
            "Falta Semanal": 0,
            "Nota P1": 0.0,
            "Nota P2": 0.0,
            "Média Final": 0.0
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
        download_name='Diario_Online_Turma_1017.xlsx'
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

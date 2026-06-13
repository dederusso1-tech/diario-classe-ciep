import os
from datetime import datetime
import pandas as pd
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, send_file

app = Flask(__name__)

# Banco de dados temporário na memória do servidor para guardar as chamadas e alunos
dados_sistema = {
    "escola": "CIEP 321",
    "turma": "1017",
    "disciplina": "Redação",
    "alunos": [],       # Lista de nomes de alunos
    "frequencia": {}    # Guardará {'data': {aluno_id: 'P' ou 'F'}}
}

HTML_COMPLETO = '''
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Portal do Docente - CIEP</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        body { background-color: #f4f6f9; font-family: system-ui, sans-serif; }
        .navbar-custom { background-color: #1a365d; color: white; }
        .card-custom { border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border: none; }
    </style>
</head>
<body>

<nav class="navbar navbar-custom p-3 mb-4">
    <div class="container d-flex justify-content-between">
        <span class="navbar-brand mb-0 h1 text-white">🍎 Diário de Frequência Inteligente</span>
        <span class="badge bg-success">Ambiente Ativo</span>
    </div>
</nav>

<div class="container">
    
    <div class="card card-custom p-4 bg-white mb-4">
        <h5 class="fw-bold text-primary mb-3">📂 Passo 1: Carregar Turma por Arquivo CSV</h5>
        <form action="/carregar-csv" method="POST" enctype="multipart/form-data" class="row g-3 align-items-end">
            <div class="col-md-3">
                <label class="form-label small fw-bold">Escola</label>
                <input type="text" name="escola" value="{{ info.escola }}" class="form-control form-control-sm" required>
            </div>
            <div class="col-md-2">
                <label class="form-label small fw-bold">Turma</label>
                <input type="text" name="turma" value="{{ info.turma }}" class="form-control form-control-sm" required>
            </div>
            <div class="col-md-3">
                <label class="form-label small fw-bold">Disciplina</label>
                <input type="text" name="disciplina" value="{{ info.disciplina }}" class="form-control form-control-sm" required>
            </div>
            <div class="col-md-4">
                <label class="form-label small fw-bold">Selecione o CSV da Escola</label>
                <div class="input-group input-group-sm">
                    <input type="file" name="arquivo_csv" class="form-control" accept=".csv" required>
                    <button type="submit" class="btn btn-primary fw-bold">🔄 Carregar Alunos</button>
                </div>
            </div>
        </form>
    </div>

    {% if info.alunos %}
    <div class="card card-custom p-4 bg-white mb-4">
        <div class="d-flex justify-content-between align-items-center mb-3 pb-2 border-bottom">
            <div>
                <h4 class="fw-bold m-0 text-dark">📋 Passo 2: Registro de Chamada Diária</h4>
                <small class="text-muted">Escola: {{ info.escola }} | Turma: {{ info.turma }} | Matéria: {{ info.disciplina }}</small>
            </div>
            
            <a href="/baixar-excel" class="btn btn-success fw-bold px-4 py-2 shadow-sm">
                📥 Passo 3: Baixar Planilha Excel Atualizada
            </a>
        </div>

        <form action="/" method="GET" class="row g-2 align-items-center mb-4 bg-light p-2 rounded">
            <div class="col-auto">
                <span class="small fw-bold text-secondary">Data do Registro:</span>
            </div>
            <div class="col-auto">
                <input type="date" name="data_filtro" value="{{ data_atual }}" class="form-control form-control-sm" onchange="this.form.submit()">
            </div>
            <div class="col-auto text-muted small">
                *(Mude a data se quiser lançar ou revisar faltas de dias anteriores)
            </div>
        </form>

        <form action="/salvar-chamada" method="POST">
            <input type="hidden" name="data_chamada" value="{{ data_atual }}">
            <div class="table-responsive">
                <table class="table table-striped table-bordered align-middle m-0">
                    <thead class="table-dark">
                        <tr>
                            <th>Nome Completo do Aluno</th>
                            <th class="text-center" style="width: 280px;">Frequência de Hoje</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for index, nome in enumerate(info.alunos) %}
                        <tr>
                            <td class="ps-3"><strong>{{ nome }}</strong></td>
                            <td class="text-center">
                                <div class="btn-group" role="group">
                                    <input type="radio" class="btn-check" name="status_{{ index }}" id="p_{{ index }}" value="P" {% if frequencia_hoje.get(index|string, 'P') == 'P' %}checked{% endif %}>
                                    <label class="btn btn-sm btn-outline-success px-3 fw-bold" for="p_{{ index }}">P (Presença)</label>

                                    <input type="radio" class="btn-check" name="status_{{ index }}" id="f_{{ index }}" value="F" {% if frequencia_hoje.get(index|string, 'P') == 'F' %}checked{% endif %}>
                                    <label class="btn btn-sm btn-outline-danger px-3 fw-bold" for="f_{{ index }}">F (Falta)</label>
                                </div>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
            <div class="text-end mt-3">
                <button type="submit" class="btn btn-primary fw-bold px-4">💾 Salvar Chamada deste Dia</button>
            </div>
        </form>
    </div>
    {% endif %}

</div>

</body>
</html>
'''

@app.route('/')
def index():
    data_filtro = request.args.get('data_filtro')
    if not data_filtro:
        data_filtro = datetime.today().strftime('%Y-%m-%d')
        
    # Busca a chamada daquele dia específico na memória
    frequencia_hoje = dados_sistema["frequencia"].get(data_filtro, {})
    
    return render_template_string(
        HTML_COMPLETO, 
        info=dados_sistema, 
        data_atual=data_filtro, 
        frequencia_hoje=frequencia_hoje,
        enumerate=enumerate
    )

@app.route('/carregar-csv', methods=['POST'])
def carregar_csv():
    global dados_sistema
    dados_sistema["escola"] = request.form.get('escola').strip().upper()
    dados_sistema["turma"] = request.form.get('turma').strip().upper()
    dados_sistema["disciplina"] = request.form.get('disciplina').strip().upper()
    file = request.files.get('arquivo_csv')

    if file:
        try:
            try:
                df = pd.read_csv(file, sep=';', encoding='utf-8')
            except:
                file.seek(0)
                df = pd.read_csv(file, sep=',', encoding='utf-8')
            
            if df.shape[1] == 0 or len(df.columns) == 0:
                file.seek(0)
                df = pd.read_csv(file, sep=';', encoding='iso-8859-1')

            coluna_nome = None
            for col in df.columns:
                if 'nome' in col.lower() or 'aluno' in col.lower() or 'estudante' in col.lower():
                    coluna_nome = col
                    break
            if not coluna_nome:
                coluna_nome = df.columns[0]

            dados_sistema["alunos"] = df[coluna_nome].dropna().astype(str).str.upper().str.strip().tolist()
            dados_sistema["frequencia"] = {} # Reseta chamadas anteriores ao carregar nova turma
        except Exception as e:
            print(f"Erro ao ler CSV: {e}")

    return redirect(url_for('index'))

@app.route('/salvar-chamada', methods=['POST'])
def salvar_chamada():
    global dados_sistema
    data_chamada = request.form.get('data_chamada')
    
    if data_chamada not in dados_sistema["frequencia"]:
        dados_sistema["frequencia"][data_chamada] = {}
        
    for index, _ in enumerate(dados_sistema["alunos"]):
        status = request.form.get(f'status_{index}', 'P')
        dados_sistema["frequencia"][data_chamada][str(index)] = status

    return redirect(url_for('index', data_filtro=data_chamada))

@app.route('/baixar-excel')
def baixar_excel():
    dados_planilha = []
    
    for index, nome in enumerate(dados_sistema["alunos"]):
        # Conta quantas faltas ('F') este aluno acumulou em todas as datas registradas
        total_faltas = 0
        for data, chamadas in dados_sistema["frequencia"].items():
            if chamadas.get(str(index)) == 'F':
                total_faltas += 1
                
        dados_planilha.append({
            "Nome Completo do Aluno": nome,
            "Faltas Acumuladas": total_faltas,
            "Média Final": 0.0
        })
        
    df = pd.DataFrame(dados_planilha)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f'TURMA_{dados_sistema["turma"]}')
    output.seek(0)
    
    nome_arquivo = f"Diario_{dados_sistema['escola']}_Turma_{dados_sistema['turma']}.xlsx".replace(" ", "_")
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=nome_arquivo
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

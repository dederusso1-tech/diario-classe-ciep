import os
import pandas as pd
from io import BytesIO
from flask import Flask, render_template_string, send_file, request, redirect, url_for

app = Flask(__name__)

# Memória temporária para guardar os alunos que o professor carregou via CSV
dados_carregados = {
    "escola": "Não informada",
    "turma": "Não informada",
    "disciplina": "Não informada",
    "alunos": []
}

HTML_GERADOR = '''
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Gerador Universal de Diários</title>
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
        <span class="navbar-brand mb-0 h1 text-white">🚀 Gerador Universal de Diários Oficiais</span>
        <span class="badge bg-success">Online</span>
    </div>
</nav>

<div class="container">
    <div class="row">
        <div class="col-md-5 mb-4">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold mb-3 text-primary">📂 Passo 1: Importar Arquivo CSV</h5>
                <p class="text-muted small">O arquivo CSV deve conter uma coluna com os nomes dos alunos.</p>
                
                <form action="/upload-csv" method="POST" enctype="multipart/form-data">
                    <div class="mb-3">
                        <label class="form-label font-weight-bold">Nome da Escola</label>
                        <input type="text" name="escola" class="form-control" placeholder="Ex: CIEP 321" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Identificação da Turma</label>
                        <input type="text" name="turma" class="form-control" placeholder="Ex: 1017 ou 2001" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Disciplina</label>
                        <input type="text" name="disciplina" class="form-control" placeholder="Ex: Redação / Matemática" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Selecione o arquivo CSV da Turma</label>
                        <input type="file" name="arquivo_csv" class="form-control" accept=".csv" required>
                    </div>
                    <button type="submit" class="btn btn-primary w-100 fw-bold">🔄 Processar e Listar Alunos</button>
                </form>
                
                {% if msg_erro %}
                    <div class="alert alert-danger mt-3 py-2 small text-center">{{ msg_erro }}</div>
                {% endif %}
            </div>
        </div>

        <div class="col-md-7">
            <div class="card card-custom p-4 bg-white">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h5 class="fw-bold m-0 text-success">📋 Dados Processados</h5>
                    {% if info.alunos %}
                        <a href="/exportar" class="btn btn-success fw-bold px-3 py-2 shadow-sm">
                            📥 Baixar Diário em Excel
                        </a>
                    {% endif %}
                </div>

                {% if not info.alunos %}
                    <div class="text-center py-5 text-muted">
                        <p class="m-0">Nenhum arquivo enviado ainda.</p>
                        <small>Preencha os dados ao lado e envie o CSV da turma para gerar a planilha.</small>
                    </div>
                {% else %}
                    <div class="alert alert-secondary py-2 mb-3 small">
                        <strong>Escola:</strong> {{ info.escola }} | <strong>Turma:</strong> {{ info.turma }} | <strong>Matéria:</strong> {{ info.disciplina }}
                    </div>
                    <p class="small text-muted mb-2">Total de alunos detectados: <strong>{{ info.alunos|length }}</strong></p>
                    <div class="table-responsive" style="max-height: 400px; overflow-y: auto;">
                        <table class="table table-striped table-sm align-middle">
                            <thead class="table-dark">
                                <tr>
                                    <th>Nome do Aluno</th>
                                </tr>
                            </thead>
                            <tbody>
                                {% for aluno in info.alunos %}
                                <tr>
                                    <td><strong>{{ aluno }}</strong></td>
                                </tr>
                                {% endfor %}
                            </tbody>
                        </table>
                    </div>
                {% endif %}
            </div>
        </div>
    </div>
</div>

</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_GERADOR, info=dados_carregados, msg_erro=None)

@app.route('/upload-csv', methods=['POST'])
def upload_csv():
    global dados_carregados
    escola = request.form.get('escola').strip()
    turma = request.form.get('turma').strip()
    disciplina = request.form.get('disciplina').strip()
    file = request.files.get('arquivo_csv')

    if not file or file.filename == '':
        return render_template_string(HTML_GERADOR, info=dados_carregados, msg_erro="Nenhum arquivo selecionado.")

    try:
        # Tenta ler o CSV de forma inteligente (testando ponto e vírgula ou vírgula)
        try:
            df = pd.read_csv(file, sep=';', encoding='utf-8')
        except:
            file.seek(0)
            df = pd.read_csv(file, sep=',', encoding='utf-8')

        # Se o CSV vier codificado em outra norma comum no Brasil (ISO-8859-1)
        if df.shape[1] == 0 or len(df.columns) == 0:
            file.seek(0)
            df = pd.read_csv(file, sep=';', encoding='iso-8859-1')

        # Procura a coluna que tem os nomes dos alunos de forma aproximada
        coluna_nome = None
        for col in df.columns:
            if 'nome' in col.lower() or 'aluno' in col.lower() or 'estudante' in col.lower():
                coluna_nome = col
                break

        # Se não achar nenhuma coluna com esses nomes, pega a primeira coluna do arquivo
        if not coluna_nome:
            coluna_nome = df.columns[0]

        # Extrai a lista de alunos limpando valores vazios
        lista_alunos = df[coluna_nome].dropna().astype(str).str.upper().str.strip().tolist()

        if not lista_alunos:
            return render_template_string(HTML_GERADOR, info=dados_carregados, msg_erro="Nenhum aluno encontrado na coluna.")

        # Guarda na memória do servidor para o download
        dados_carregados = {
            "escola": escola,
            "turma": turma,
            "disciplina": disciplina,
            "alunos": lista_alunos
        }
        return redirect(url_for('index'))

    except Exception as e:
        return render_template_string(HTML_GERADOR, info=dados_carregados, msg_erro=f"Erro ao ler o CSV: {str(e)}")

@app.route('/exportar')
def exportar():
    if not dados_carregados["alunos"]:
        return redirect(url_for('index'))
        
    dados_planilha = []
    for aluno in dados_carregados["alunos"]:
        dados_planilha.append({
            "Nome Completo": aluno,
            "Faltas": 0,
            "Nota 1": 0.0,
            "Nota 2": 0.0,
            "Média Final": 0.0
        })
        
    df = pd.DataFrame(dados_planilha)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f'TURMA_{dados_carregados["turma"]}')
    output.seek(0)
    
    nome_arquivo = f"Diario_{dados_carregados['escola']}_Turma_{dados_carregados['turma']}.xlsx".replace(" ", "_")
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=nome_arquivo
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port) os
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

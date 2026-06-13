import os
import re
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Banco de dados local permanente no Render
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'diario_ciep_colunas.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

# --- MODELOS DO BANCO DE DADOS ---
class Turma(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    escola = db.Column(db.String(150), nullable=False)
    nome_turma = db.Column(db.String(50), nullable=False)
    disciplina = db.Column(db.String(100), nullable=False)
    alunos = db.relationship('Aluno', backref='turma', lazy=True, cascade="all, delete-orphan")

class Aluno(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    matricula = db.Column(db.String(50), nullable=True)
    nome = db.Column(db.String(200), nullable=False)
    situacao = db.Column(db.String(50), nullable=True)
    turma_id = db.Column(db.Integer, db.ForeignKey('turma.id'), nullable=False)
    presencas = db.relationship('Presenca', backref='aluno', lazy=True, cascade="all, delete-orphan")

class Presenca(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data = db.Column(db.String(10), nullable=False)
    status = db.Column(db.String(1), nullable=False) # 'P' ou 'F'
    aluno_id = db.Column(db.Integer, db.ForeignKey('aluno.id'), nullable=False)

# --- INTERFACE HTML REFORMULADA COM 4 COLUNAS ---
HTML_COMPLETO = '''
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <title>Portal do Docente - CIEP 205</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    <style>
        body { background-color: #f4f6f9; font-family: system-ui, sans-serif; }
        .navbar-custom { background-color: #1a365d; color: white; }
        .card-custom { border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); border: none; }
        .table th { background-color: #2d3748 !important; color: white !important; }
    </style>
</head>
<body>

<nav class="navbar navbar-custom p-3 mb-4">
    <div class="container d-flex justify-content-between">
        <span class="navbar-brand mb-0 h1 text-white">🍎 Diário de Frequência Inteligente - CIEP 205</span>
        <a href="/" class="btn btn-sm btn-outline-light fw-bold">🏠 Voltar ao Menu de Turmas</a>
    </div>
</nav>

<div class="container">
    {% if tela == 'inicial' %}
    <div class="row">
        <div class="col-md-5 mb-4">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-primary mb-3">📂 Importar Nova Turma (CSV do Sistema)</h5>
                <form action="/carregar-csv" method="POST" enctype="multipart/form-data">
                    <div class="mb-2">
                        <label class="form-label small fw-bold">Unidade Escolar</label>
                        <input type="text" name="escola" class="form-control form-control-sm" value="CIEP 205 FREI AGOSTINHO FÍNCIAS" required>
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold">Código da Turma</label>
                        <input type="text" name="turma" class="form-control form-control-sm" placeholder="Ex: 1005" required>
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold">Disciplina</label>
                        <input type="text" name="disciplina" class="form-control form-control-sm" placeholder="Ex: LÍNGUA PORTUGUESA" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold">Selecione o Arquivo CSV do Sistema</label>
                        <input type="file" name="arquivo_csv" class="form-control form-control-sm" accept=".csv" required>
                    </div>
                    <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold">🔄 Organizar em Colunas e Salvar</button>
                </form>
            </div>
        </div>

        <div class="col-md-7">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-success mb-3">📚 Suas Turmas Salvas na Nuvem</h5>
                {% if not turmas %}
                    <p class="text-muted small">Nenhuma turma processada ainda. Faça o upload do arquivo ao lado.</p>
                {% else %}
                    <div class="list-group">
                        {% for t in turmas %}
                            <div class="list-group-item d-flex justify-content-between align-items-center mb-2 rounded border">
                                <div>
                                    <h6 class="fw-bold m-0">Turma {{ t.nome_turma }} - {{ t.disciplina }}</h6>
                                    <small class="text-muted">{{ t.escola }}</small>
                                </div>
                                <div>
                                    <a href="/chamada/{{ t.id }}" class="btn btn-sm btn-success fw-bold">📅 Chamada</a>
                                    <a href="/excluir-turma/{{ t.id }}" class="btn btn-sm btn-outline-danger" onclick="return confirm('Apagar esta turma permanentemente?')">🗑️</a>
                                </div>
                            </div>
                        {% endfor %}
                    </div>
                {% endif %}
            </div>
        </div>
    </div>

    {% elif tela == 'chamada' %}
    <div class="card card-custom p-4 bg-white mb-4">
        <div class="d-flex justify-content-between align-items-center mb-3 pb-2 border-bottom">
            <div>
                <h4 class="fw-bold m-0 text-dark">📋 Frequência Escolar Organizada</h4>
                <small class="text-muted">Componente: {{ turma.disciplina }} | Turma: {{ turma.nome_turma }}</small>
            </div>
            <a href="/baixar-excel/{{ turma.id }}" class="btn btn-success btn-sm fw-bold px-4 shadow-sm">
                📥 Exportar Diário Caderneta (Excel)
            </a>
        </div>

        <form action="/chamada/{{ turma.id }}" method="GET" class="row g-2 align-items-center mb-4 bg-light p-2 rounded">
            <div class="col-auto">
                <span class="small fw-bold text-secondary">Data Letiva:</span>
            </div>
            <div class="col-auto">
                <input type="date" name="data_filtro" value="{{ data_atual }}" class="form-control form-control-sm" onchange="this.form.submit()">
            </div>
        </form>

        <form action="/salvar-chamada/{{ turma.id }}" method="POST">
            <input type="hidden" name="data_chamada" value="{{ data_atual }}">
            <div class="table-responsive">
                <table class="table table-striped table-bordered align-middle m-0 table-sm">
                    <thead>
                        <tr>
                            <th class="ps-2">Nº Matrícula</th>
                            <th>Nome Completo do Aluno</th>
                            <th class="text-center">Situação</th>
                            <th class="text-center" style="width: 250px;">Lançamento de Hoje</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for aluno in alunos_info %}
                        <tr>
                            <td class="ps-2 text-secondary small"><code>{{ aluno.matricula }}</code></td>
                            <td><strong class="text-dark">{{ aluno.nome }}</strong></td>
                            <td class="text-center">
                                <span class="badge bg-success-subtle text-success border border-success-subtle rounded-pill small px-2">{{ aluno.situacao }}</span>
                            </td>
                            <td class="text-center">
                                <div class="btn-group" role="group">
                                    <input type="radio" class="btn-check" name="status_{{ aluno.id }}" id="p_{{ aluno.id }}" value="P" {% if aluno.status_hoje == 'P' %}checked{% endif %}>
                                    <label class="btn btn-xs btn-outline-success px-2 fw-bold small" for="p_{{ aluno.id }}">P</label>

                                    <input type="radio" class="btn-check" name="status_{{ aluno.id }}" id="f_{{ aluno.id }}" value="F" {% if aluno.status_hoje == 'F' %}checked{% endif %}>
                                    <label class="btn btn-xs btn-outline-danger px-2 fw-bold small" for="f_{{ aluno.id }}">F</label>
                                </div>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
            <div class="text-end mt-3">
                <button type="submit" class="btn btn-primary fw-bold px-4">💾 Gravar Chamada deste Dia</button>
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
    turmas = Turma.query.all()
    return render_template_string(HTML_COMPLETO, tela='inicial', turmas=turmas)

@app.route('/carregar-csv', methods=['POST'])
def carregar_csv():
    escola = request.form.get('escola').strip().upper()
    nome_turma = request.form.get('turma').strip().upper()
    disciplina = request.form.get('disciplina').strip().upper()
    file = request.files.get('arquivo_csv')

    if file:
        try:
            conteudo_bruto = file.read().decode('utf-8', errors='ignore').strip()
            
            # Divide o bloco contínuo localizando os anos de matrícula padrão
            blocos_alunos = re.split(r'(?=2024\d{11}|2025\d{11}|2026\d{11})', conteudo_bruto)
            
            nova_turma = Turma(escola=escola, nome_turma=nome_turma, disciplina=disciplina)
            db.session.add(nova_turma)
            db.session.commit()

            for bloco in blocos_alunos:
                bloco = bloco.strip()
                # Cadastra apenas quem está MATRICULADO e descarta os CANCELADOS
                if "MATRICULADO" in bloco and "CANCELADO" not in bloco:
                    partes = [p.strip() for p in bloco.split(',')]
                    
                    if len(partes) >= 4:
                        mat_candidata = partes[0]
                        sit_candidata = "MATRICULADO"
                        nome_candidato = ""
                        
                        # Captura o nome limpando o resto das vírgulas internas
                        for parte in partes:
                            if len(parte) > 10 and not any(c.isdigit() for c in parte) and "TRIMESTRE" not in parte:
                                nome_candidato = parte.upper()
                                break
                        
                        if nome_candidato:
                            aluno = Aluno(
                                matricula=mat_candidata,
                                nome=nome_candidato,
                                situacao=sit_candidata,
                                turma_id=nova_turma.id
                            )
                            db.session.add(aluno)
            db.session.commit()
                
        except Exception as e:
            print(f"Erro no fatiador: {e}")

    return redirect(url_for('index'))

@app.route('/chamada/<int:turma_id>')
def chamada(turma_id):
    turma = Turma.query.get_or_404(turma_id)
    data_filtro = request.args.get('data_filtro')
    if not data_filtro:
        data_filtro = datetime.today().strftime('%Y-%m-%d')

    alunos = Aluno.query.filter_by(turma_id=turma.id).order_by(Aluno.nome).all()
    alunos_info = []
    
    for a in alunos:
        reg = Presenca.query.filter_by(aluno_id=a.id, data=data_filtro).first()
        status_hoje = reg.status if reg else 'P'
        alunos_info.append({
            "id": a.id, 
            "matricula": a.matricula,
            "nome": a.nome, 
            "situacao": a.situacao,
            "status_hoje": status_hoje
        })

    return render_template_string(HTML_COMPLETO, tela='chamada', turma=turma, alunos_info=alunos_info, data_atual=data_filtro)

@app.route('/salvar-chamada/<int:turma_id>', methods=['POST'])
def salvar_chamada(turma_id):
    data_chamada = request.form.get('data_chamada')
    turma = Turma.query.get_or_404(turma_id)
    alunos = Aluno.query.filter_by(turma_id=turma.id).all()

    for a in alunos:
        status = request.form.get(f'status_{a.id}', 'P')
        reg = Presenca.query.filter_

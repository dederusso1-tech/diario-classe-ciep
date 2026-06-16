import os
import re
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
import openpyxl

app = Flask(__name__)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
# Banco definitivo unificado de produção
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'diario_ciep_producao_final.db')
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
    num_chamada = db.Column(db.String(10), nullable=True)
    matricula = db.Column(db.String(50), nullable=True)
    nome = db.Column(db.String(200), nullable=False)
    situacao = db.Column(db.String(50), nullable=True)
    nota1 = db.Column(db.Float, default=0.0)
    nota2 = db.Column(db.Float, default=0.0)
    turma_id = db.Column(db.Integer, db.ForeignKey('turma.id'), nullable=False)
    presencas = db.relationship('Presenca', backref='aluno', lazy=True, cascade="all, delete-orphan")

class Presenca(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data = db.Column(db.String(10), nullable=False)
    status = db.Column(db.String(1), nullable=False)
    aluno_id = db.Column(db.Integer, db.ForeignKey('aluno.id'), nullable=False)

with app.app_context():
    db.create_all()

# --- INTERFACE HTML ---
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
        .table th { background-color: #1a365d !important; color: white !important; padding: 12px; text-align: center; }
        .table td { vertical-align: middle; padding: 10px; }
        .input-nota { width: 75px; text-align: center; font-weight: bold; }
    </style>
</head>
<body>
<nav class="navbar navbar-custom p-3 mb-4">
    <div class="container d-flex justify-content-between">
        <span class="navbar-brand mb-0 h1 text-white">🍎 Diário de Frequência Inteligente</span>
        <a href="/" class="btn btn-sm btn-outline-light fw-bold">🏠 Voltar ao Menu</a>
    </div>
</nav>
<div class="container">
    {% if tela == 'inicial' %}
    <div class="row">
        <div class="col-md-5 mb-4">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-primary mb-3">📂 Carregar Novo Diário (CSV)</h5>
                <form action="/carregar-csv" method="POST" enctype="multipart/form-data">
                    <div class="mb-3">
                        <label class="form-label small fw-bold">Selecione o Arquivo CSV</label>
                        <input type="file" name="arquivo_csv" class="form-control form-control-sm" accept=".csv" required>
                    </div>
                    <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold">🔄 Criar Diário na Nuvem</button>
                </form>
            </div>
        </div>
        <div class="col-md-7">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-success mb-3">📚 Suas Turmas Salvas na Nuvem</h5>
                {% if not turmas %}
                    <p class="text-muted small">Nenhuma turma cadastrada ainda.</p>
                {% else %}
                    <div class="list-group">
                        {% for t in turmas %}
                            <div class="list-group-item d-flex justify-content-between align-items-center mb-2 rounded border">
                                <div>
                                    <h6 class="fw-bold m-0">Turma {{ t.nome_turma }} - {{ t.disciplina }}</h6>
                                    <small class="text-muted">{{ t.escola }}</small>
                                </div>
                                <div>
                                    <a href="/chamada/{{ t.id }}" class="btn btn-sm btn-success fw-bold">📅 Caderneta</a>
                                    <a href="/excluir-turma/{{ t.id }}" class="btn btn-sm btn-outline-danger">🗑️</a>
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
                <h4 class="fw-bold m-0 text-dark">📋 Fechamento de Notas e Frequência</h4>
                <small class="text-muted">{{ turma.escola }} | {{ turma.disciplina }} | Turma: {{ turma.nome_turma }}</small>
                <div class="mt-2"><span class="badge bg-primary fs-6">📅 Data Selecionada: 16/06/2026</span></div>
            </div>
            <a href="/baixar-excel/{{ turma.id }}" class="btn btn-success btn-sm fw-bold px-4 shadow-sm">📥 Exportar Planilha de Impressão</a>
        </div>
        <form action="/salvar-dados/{{ turma.id }}" method="POST">
            <input type="hidden" name="data_chamada" value="{{ data_atual }}">
            <div class="table-responsive">
                <table class="table table-striped table-bordered align-middle m-0 table-sm">
                    <thead>
                        <tr>
                            <th style="width: 50px;">Nº</th>
                            <th>Nome Completo do Aluno</th>
                            <th style="width: 80px;">Faltas</th>
                            <th style="width: 95px;">Nota 1</th>
                            <th style="width: 95px;">Nota 2</th>
                            <th style="width: 95px;">Média</th>
                            <th style="width: 130px;">Chamada de Hoje (16/06)</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for aluno in alunos_info %}
                        <tr>
                            <td class="text-center text-muted small">{{ aluno.num_chamada }}</td>
                            <td><strong class="text-dark">{{ aluno.nome }}</strong><br><small class="text-muted">{{ aluno.matricula }}</small></td>
                            <td class="text-center text-danger fw-bold">{{ aluno.total_faltas }}</td>
                            <td class="text-center">
                                <input type="number" step="0.1" min="0" max="10" name="nota1_{{ aluno.id }}" class="form-control form-control-sm input-nota text-primary" value="{{ aluno.nota1 }}">
                            </td>
                            <td class="text-center">
                                <input type="number" step="0.1" min="0" max="10" name="nota2_{{ aluno.id }}" class="form-control form-control-sm input-nota text-primary" value="{{ aluno.nota2 }}">
                            </td>
                            <td class="text-center fw-bold {% if aluno.media >= 6.0 %}text-success{% else %}text-danger{% endif %}">
                                {{ "%.1f"|format(aluno.media) }}
                            </td>
                            <td class="text-center">
                                <div class="btn-group">
                                    <input type="radio" class="btn-check" name="status_{{ aluno.id }}" id="p_{{ aluno.id }}" value="P" {% if aluno.status_hoje == 'P' %}checked{% endif %}>
                                    <label class="btn btn-xs btn-outline-success px-2 fw-bold" for="p_{{ aluno.id }}">P</label>
                                    <input type="radio" class="btn-check" name="status_{{ aluno.id }}" id="f_{{ aluno.id }}" value="F" {% if aluno.status_hoje == 'F' %}checked{% endif %}>
                                    <label class="btn btn-xs btn-outline-danger px-2 fw-bold" for="f_{{ aluno.id }}">F</label>
                                </div>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
            <button type="submit" class="btn btn-primary fw-bold px-5 mt-3 shadow">💾 Gravar Chamada e Notas (Dia 16/06)</button>
        </form>
    </div>
    {% endif %}
</div>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_COMPLETO, tela='inicial', turmas=Turma.query.all())

@app.route('/carregar-csv', methods=['POST'])
def carregar_csv():
    file = request.files.get('arquivo_csv')
    if file:
        try:
            texto_bruto = file.read().decode('utf-8', errors='ignore')
            linhas = texto_bruto.replace('\r', '\n').split('\n')
            
            escola_auto =

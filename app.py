import os
from datetime import datetime
from io import BytesIO
import pandas as pd
from flask import Flask, render_template_string, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
import openpyxl

app = Flask(__name__)

# Banco de dados estável para o ambiente do Render
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'diario_ciep_limpo_v2.db')
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
    turma_id = db.Column(db.Integer, db.ForeignKey('turma.id'), nullable=False)
    presencas = db.relationship('Presenca', backref='aluno', lazy=True, cascade="all, delete-orphan")

class Presenca(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data = db.Column(db.String(10), nullable=False)
    status = db.Column(db.String(1), nullable=False)
    aluno_id = db.Column(db.Integer, db.ForeignKey('aluno.id'), nullable=False)

# --- INTERFACE HTML LIMPA E ESPAÇADA ---
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
        .table th { background-color: #1a365d !important; color: white !important; padding: 10px; }
        .table td { vertical-align: middle; }
    </style>
</head>
<body>
<nav class="navbar navbar-custom p-3 mb-4">
    <div class="container d-flex justify-content-between">
        <span class="navbar-brand mb-0 h1 text-white">🍎 Diário de Frequência Inteligente - CIEP 205</span>
        <a href="/" class="btn btn-sm btn-outline-light fw-bold">🏠 Voltar ao Menu</a>
    </div>
</nav>
<div class="container">
    {% if tela == 'inicial' %}
    <div class="row">
        <div class="col-md-5 mb-4">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-primary mb-3">📂 Carregar Novo CSV de Diário</h5>
                <form action="/carregar-csv" method="POST" enctype="multipart/form-data">
                    <div class="mb-2">
                        <label class="form-label small fw-bold">Unidade Escolar</label>
                        <input type="text" name="escola" class="form-control form-control-sm" value="CIEP 205 FREI AGOSTINHO FÍNCIAS" required>
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold">Código da Turma</label>
                        <input type="text" name="turma" class="form-control form-control-sm" placeholder="Ex: 1017" required>
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold">Componente Curricular</label>
                        <input type="text" name="disciplina" class="form-control form-control-sm" placeholder="Ex: LÍNGUA PORTUGUESA" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold">Selecione o Arquivo CSV do Sistema</label>
                        <input type="file" name="arquivo_csv" class="form-control form-control-sm" accept=".csv" required>
                    </div>
                    <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold">🔄 Remover Sujeira e Alinhar Colunas</button>
                </form>
            </div>
        </div>
        <div class="col-md-7">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-success mb-3">📚 Suas Turmas Ativas</h5>
                {% if not turmas %}
                    <p class="text-muted small">Nenhuma turma cadastrada ainda.</p>
                {% else %}
                    <div class="list-group">
                        {% for t in turmas %}
                            <div class="list-group-item d-flex justify-content-between align-items-center mb-2 rounded border">
                                <div><h6 class="fw-bold m-0">Turma {{ t.nome_turma }} - {{ t.disciplina }}</h6></div>
                                <div>
                                    <a href="/chamada/{{ t.id }}" class="btn btn-sm btn-success fw-bold">📅 Chamada</a>
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
                <h4 class="fw-bold m-0 text-dark">📋 Lista de Chamada Limpa e Alinhada</h4>
                <small class="text-muted">{{ turma.disciplina }} | Turma: {{ turma.nome_turma }}</small>
            </div>
            <a href="/baixar-excel/{{ turma.id }}" class="btn btn-success btn-sm fw-bold px-4 shadow-sm">📥 Exportar para Excel</a>
        </div>
        <form action="/salvar-chamada/{{ turma.id }}" method="POST">
            <input type="hidden" name="data_chamada" value="{{ data_atual }}">
            <div class="table-responsive">
                <table class="table table-striped table-bordered align-middle m-0 table-sm">
                    <thead>
                        <tr>
                            <th class="text-center" style="width: 60px;">Nº</th>
                            <th style="width: 160px;">Matrícula</th>
                            <th>Nome Completo do Aluno</th>
                            <th class="text-center" style="width: 140px;">Situação</th>
                            <th class="text-center" style="width: 150px;">Frequência de Hoje</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for aluno in alunos_info %}
                        <tr>
                            <td class="text-center text-muted small">{{ aluno.num_chamada }}</td>
                            <td class="text-secondary small"><code>{{ aluno.matricula }}</code></td>
                            <td><span class="text-dark fw-semibold">{{ aluno.nome }}</span></td>
                            <td class="text-center"><span class="badge bg-success text-white small px-2">{{ aluno.situacao }}</span></td>
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
            <button type="submit" class="btn btn-primary fw-bold px-4 mt-3">💾 Gravar Chamada</button>
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
    escola = request.form.get('escola').strip().upper()
    nome_turma = request.form.get('turma').strip().upper()
    disciplina

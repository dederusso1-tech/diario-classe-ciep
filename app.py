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

# Banco de dados estável no Render
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'diario_ciep_limpo.db')
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
    turma_id = db.Column(db.Integer, db.ForeignKey('turma.id'), nullable=False)
    presencas = db.relationship('Presenca', backref='aluno', lazy=True, cascade="all, delete-orphan")

class Presenca(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    data = db.Column(db.String(10), nullable=False)
    status = db.Column(db.String(1), nullable=False) # 'P' ou 'F'
    aluno_id = db.Column(db.Integer, db.ForeignKey('aluno.id'), nullable=False)

# --- INTERFACE HTML ---
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
        .table th { background-color: #1a365d !important; color: white !important; }
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
                    <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold">🔄 Extrair Nomes e Salvar Turma</button>
                </form>
            </div>
        </div>

        <div class="col-md-7">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-success mb-3">📚 Suas Turmas Salvas na Nuvem</h5>
                {% if not turmas %}
                    <p class="text-muted small">Nenhuma turma ativa armazenada na nuvem. Use o formulário para carregar o CSV.</p>
                {% else %}
                    <div class="list-group">
                        {% for t in turmas %}
                            <div class="list-group-item d-flex justify-content-between align-items-center mb-2 rounded border">
                                <div>
                                    <h6 class="fw-bold m-0">Turma {{ t.nome_turma }} - {{ t.disciplina }}</h6>
                                    <small class="text-muted">{{ t.escola }}</small>
                                </div>
                                <div>
                                    <a href="/chamada/{{ t.id }}" class="btn btn-sm btn-success fw-bold">📅 Fazer Chamada</a>
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
                <h4 class="fw-bold m-0 text-dark">📋 Lista de Chamada de Alunos Ativos</h4>
                <small class="text-muted">Componente: {{ turma.disciplina }} | Turma: {{ turma.nome_turma }}</small>
            </div>
            <a href="/baixar-excel/{{ turma.id }}" class="btn btn-success fw-bold px-4 shadow-sm">
                📥 Baixar Caderneta Pronta (Excel)
            </a>
        </div>

        <form action="/chamada/{{ turma.id }}" method="GET" class="row g-2 align-items-center mb-4 bg-light p-2 rounded">
            <div class="col-auto">
                <span class="small fw-bold text-secondary">Data do Registro:</span>
            </div>
            <div class="col-auto">
                <input type="date" name="data_filtro" value="{{ data_atual }}" class="form-control form-control-sm" onchange="this.form.submit()">
            </div>
        </form>

        <form action="/salvar-chamada/{{ turma.id }}" method="POST">
            <input type="hidden" name="data_chamada" value="{{ data_atual }}">
            <div class="table-responsive">

import os
import pandas as pd
from io import BytesIO
from flask import Flask, render_template_string, send_file, request, redirect, url_for

app = Flask(__name__)

# Banco de dados temporário na memória do servidor
banco_dados = {
    "professores": {},  # Guardará {usuario: senha}
    "diarios": {}       # Guardará {id_diario: {turma, disciplina, alunos}}
}
id_diario_control = 1

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
        {% if professor_atual %}
            <span class="text-white-50">Professor: <strong>{{ professor_atual }}</strong> | <a href="/" class="text-warning text-decoration-none">Sair</a></span>
        {% endif %}
    </div>
</nav>

<div class="container">
    {% if tela == 'cadastro' %}
    <div class="row justify-content-center">
        <div class="col-md-5">
            <div class="card card-custom p-4 bg-white">
                <h4 class="fw-bold text-center text-dark mb-4">🔐 Acesso ao Sistema</h4>
                
                <form action="/login-cadastro" method="POST">
                    <div class="mb-3">
                        <label class="form-label">Usuário / E-mail</label>
                        <input type="text" name="usuario" class="form-control" required placeholder="ex: andre.brito">
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Senha</label>
                        <input type="password" name="senha" class="form-control" required placeholder="Sua senha">
                    </div>
                    <div class="d-grid gap-2">
                        <button type="submit" name="acao" value="login" class="btn btn-primary fw-bold">Entrar no Sistema</button>
                        <button type="submit" name="acao" value="cadastro" class="btn btn-success fw-bold">Criar Nova Conta de Professor</button>
                    </div>
                </form>
                {% if msg %}
                    <div class="alert alert-info mt-3 text-center py-2 small">{{ msg }}</div>
                {% endif %}
            </div>
        </div>
    </div>

    {% elif tela == 'dashboard' %}
    <div class="row">
        <div class="col-md-4 mb-4">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold mb-3">🆕 Criar Novo Diário</h5>
                <form action="/criar-diario" method="POST">
                    <input type="hidden" name="professor" value="{{ professor_atual }}">
                    <div class="mb-3">
                        <label class="form-label">Número da Turma</label>
                        <input type="text" name="turma" class="form-control" required placeholder="Ex: 1017">
                    </div>
                    <div class="mb-3">
                        <label class="form-label">Disciplina</label>
                        <input type="text" name="disciplina" class="form-control" required placeholder="Ex: Redação">
                    </div>
                    <button type="submit" class="btn btn-success w-100 fw-bold">+ Inicializar Diário</button>
                </form>
            </div>
        </div>
        
        <div class="col-md-8">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold mb-3">📚 Meus Diários Ativos</h5>
                {% if n_diarios == 0 %}
                    <p class="text-muted">Você ainda não criou nenhum diário. Use o formulário ao lado para começar!</p>
                {% else %}
                    <div class="list-group">
                        {% for id, info in diarios.items() %}
                            <a href="/diario/{{ id }}?usuario={{ professor_atual }}" class="list-group-item list-group-item-action d-flex justify-content-between align-items-center">
                                <div>
                                    <h6 class="fw-bold mb-1">Turma {{ info.turma }}</h6>
                                    <small class="text-muted">Disciplina: {{ info.disciplina }}</small>
                                </div>
                                <span class="badge bg-primary rounded-pill">Abrir Diário →</span>
                            </a>
                        {% endfor %}
                    </div>
                {% endif %}
            </div>
        </div>
    </div>

    {% elif tela == 'ver_diario' %}
    <div class="card card-custom p-4 bg-white mb-4">
        <p class="text-muted small"><a href="/painel?usuario={{ professor_atual }}" class="text-decoration-none">← Voltar para o Painel</a></p>
        <div class="d-flex justify-content-between align-items-center mb-4">
            <div>
                <h3 class="fw-bold m-0">Turma {{ diario_info.turma }}</h3>
                <span class="text-muted">Matéria: {{ diario_info.disciplina }}</span>
            </div>
            <a href="/exportar-excel/{{ id_diario }}" class="btn btn-success fw-bold px-4 py-2 shadow-sm">
                📥 Baixar Planilha Excel
            </a>
        </div>

        <h5 class="fw-bold mb-3 text-secondary">📋 Alunos Cadastrados</h5>
        <form action="/adicionar-aluno/{{ id_diario }}" method="POST" class="row g-3 mb-4">
            <input type="hidden" name="professor" value="{{ professor_atual }}">
            <div class="col-md-9">
                <input type="text" name="nome_aluno" class="form-control" required placeholder="Digite o nome completo do aluno">
            </div>
            <div class="col-md-3">
                <button type="submit" class="btn btn-primary w-100 fw-bold">+ Matricular</button>
            </div>
        </form>

        <div class="table-responsive">
            <table class="table table-striped align-middle">
                <thead>
                    <tr>
                        <th>Nome Completo</th>
                        <th class="text-center" style="width: 150px;">Status</th>
                    </tr>
                </thead>
                <tbody>
                    {% for aluno in diario_info.alunos %}
                    <tr>
                        <td><strong>{{ aluno }}</strong></td>
                        <td class="text-center"><span class="badge bg-success">Ativo</span></td>
                    </tr>
                    {% else %}
                    <tr>
                        <td colspan="2" class="text-muted text-center py-3">Nenhum aluno matriculado ainda. Digite acima para adicionar.</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
    </div>
    {% endif %}
</div>

</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_SISTEMA, tela='cadastro', msg=None, professor_atual=None)

@app.route('/login-cadastro', methods=['POST'])
def login_cadastro():
    usuario = request.form.get('usuario').strip()
    senha = request.form.get('senha').strip()
    acao = request.form.get('acao')
    
    if not usuario or not

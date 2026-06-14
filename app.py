import os
import re
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

app = Flask(__name__)

# Banco de dados local estável no servidor do Render
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'diario_ciep_limpeza_total.db')
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
    status = db.Column(db.String(1), nullable=False)
    aluno_id = db.Column(db.Integer, db.ForeignKey('aluno.id'), nullable=False)

# --- INTERFACE HTML VISUAL ---
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
        <a href="/" class="btn btn-sm btn-outline-light fw-bold">🏠 Voltar ao Menu</a>
    </div>
</nav>
<div class="container">
    {% if tela == 'inicial' %}
    <div class="row">
        <div class="col-md-5 mb-4">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-primary mb-3">📂 Filtrar e Carregar CSV do Sistema</h5>
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
                        <input type="text" name="disciplina" class="form-control form-control-sm" placeholder="Ex: AS LINGUAGENS NA TECNOLOGIA" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold">Selecione o Arquivo CSV Em Linha Única</label>
                        <input type="file" name="arquivo_csv" class="form-control form-control-sm" accept=".csv" required>
                    </div>
                    <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold">🔄 Faxina Geral e Limpar Lista</button>
                </form>
            </div>
        </div>
        <div class="col-md-7">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-success mb-3">📚 Suas Turmas Salvas na Nuvem</h5>
                {% if not turmas %}
                    <p class="text-muted small">Nenhuma turma processada ainda nesta sessão.</p>
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
                <h4 class="fw-bold m-0 text-dark">📋 Frequência Escolar Organizada (Apenas Ativos)</h4>
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
                            <th class="ps-2">Matrícula</th>
                            <th>Nome Completo do Aluno</th>
                            <th class="text-center">Situação</th>
                            <th class="text-center" style="width: 180px;">Frequência de Hoje</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for aluno in alunos_info %}
                        <tr>
                            <td class="ps-2 text-secondary small"><code>{{ aluno.matricula }}</code></td>
                            <td><strong class="text-dark">{{ aluno.nome }}</strong></td>
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
    disciplina = request.form.get('disciplina').strip().upper()
    file = request.files.get('arquivo_csv')

    if file:
        try:
            conteudo = file.read().decode('utf-8', errors='ignore').strip()
            
            # Divide o bloco contínuo localizando o padrão de matrículas do estado (números longos)
            blocos = re.split(r'(?=2024\d{11}|2025\d{11}|2026\d{11})', conteudo)
            
            nova_turma = Turma(escola=escola, nome_turma=nome_turma, disciplina=disciplina)
            db.session.add(nova_turma)
            db.session.commit()

            lista_nomes_inseridos = set()

            for bloco in blocos:
                bloco = bloco.strip()
                
                # SUPER FILTRO: Ignora metadados do professor, cabeçalhos repetidos e alunos cancelados
                if "ANDRE CAMARGO" in bloco or "CIEP 205" in bloco or "NUM_CHAMADA" in bloco or "TEXTBOX" in bloco or "CANCELADO" in bloco:
                    continue
                
                if "MATRICULADO" in bloco:
                    partes = [p.strip() for p in bloco.split(',')]
                    if len(partes) >= 4:
                        mat = partes[0]
                        nome = ""
                        
                        # Extrai cirurgicamente apenas o nome do aluno ativo
                        for parte in partes:
                            if len(parte) > 8 and not any(c.isdigit() for c in parte) and "TRIMESTRE" not in parte and "MATRICULADO" not in parte:
                                nome = parte.upper()
                                break
                        
                        # Evita duplicar alunos e garante que pegou um nome válido
                        if nome and nome not in lista_nomes_inseridos:
                            lista_nomes_inseridos.add(nome)
                            db.session.add(Aluno(matricula=mat, nome=nome, situacao="MATRICULADO", turma_id=nova_turma.id))
            
            db.session.commit()
        except Exception as e:
            print(f"Erro no processamento da limpeza: {e}")
            
    return redirect(url_for('index'))

@app.route('/chamada/<int:turma_id>')
def chamada(turma_id):
    turma = Turma.query.get_or_404(turma_id)
    data_filtro = datetime.today().strftime('%Y-%m-%d')
    alunos = Aluno.query.filter_by(turma_id=turma.id).order_by(Aluno.nome).all()
    alunos_info = [{"id": a.id, "matricula": a.matricula, "nome": a.nome, "situacao": a.situacao, "status_hoje": (Presenca.query.filter_by(aluno_id=a.id, data=data_filtro).first().status if Presenca.query.filter_by(aluno_id=a.id, data=data_filtro).first() else 'P')} for a in alunos]
    return render_template_string(HTML_COMPLETO, tela='chamada', turma=turma, alunos_info=alunos_info, data_atual=data_filtro)

@app.route('/salvar-chamada/<int:turma_id>', methods=['POST'])
def salvar_chamada(turma_id):
    data_chamada = request.form.get('data_chamada')
    alunos = Aluno.query.filter_by(turma_id=turma_id).all()
    for a in alunos:
        status = request.form.get(f'status_{a.id}', 'P')
        reg = Presenca.query.filter_by(aluno_id=a.id, data=data_chamada).first()
        if reg: reg.status = status
        else: db.session.add(Presenca(data=data_chamada, status=status, aluno_id=a.id))
    db.session.commit()
    return redirect(url_for('chamada', turma_id=turma_id))

@app.route('/excluir-turma/<int:turma_id>')
def excluir_turma(turma_id):
    turma = Turma.query.get(turma_id)
    if r:= turma: db.session.delete(r); db.session.commit()
    return redirect(url_for('index'))

@app.route('/baixar-excel/<int:turma_id>')
def baixar_excel(turma_id):
    turma = Turma.query.get_or_404(turma_id)
    alunos = Aluno.query.filter_by(turma_id=turma.id).order_by(Aluno.nome).all()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DIARIO"
    ws.views.sheetView[0].showGridLines = True
    
    headers = ["Matrícula", "Nome Completo do Aluno", "Situação"]
    for col, h in enumerate(headers, 1): ws.cell(row=1, column=col, value=h)
    
    for idx, a in enumerate(alunos, 2):
        ws.cell(row=idx, column=1, value=a.matricula)
        ws.cell(row=idx, column=2, value=a.nome)
        ws.cell(row=idx, column=3, value=a.situacao)
        
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f"Diario_Limpo_Turma_{turma.nome_turma}.xlsx")

if __name__ == '__main__':
    with app.app_context(): db.create_all()
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

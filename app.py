import os
import re
from datetime import datetime
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
import openpyxl

app = Flask(__name__)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'diario_ciep_estavel_final.db')
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
        .table th { background-color: #1a365d !important; color: white !important; padding: 12px; }
        .table td { vertical-align: middle; padding: 10px; }
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
                <h5 class="fw-bold text-primary mb-3">📂 Carregar Diário Oficial (CSV)</h5>
                <form action="/carregar-csv" method="POST" enctype="multipart/form-data">
                    <div class="mb-3">
                        <label class="form-label small fw-bold">Selecione o Arquivo CSV do Diário</label>
                        <input type="file" name="arquivo_csv" class="form-control form-control-sm" accept=".csv" required>
                    </div>
                    <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold">🔄 Processar Arquivo Oficial</button>
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
                <h4 class="fw-bold m-0 text-dark">📋 Lista de Chamada Alinhada (Apenas Ativos)</h4>
                <small class="text-muted">{{ turma.escola }} | {{ turma.disciplina }} | Turma: {{ turma.nome_turma }}</small>
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
    file = request.files.get('arquivo_csv')
    if file:
        try:
            texto_bruto = file.read().decode('utf-8', errors='ignore')
            linhas = texto_bruto.replace('\r', '\n').split('\n')
            
            escola_auto = "CIEP 321 DOUTOR ULYSSES GUIMARAES"
            turma_auto = "1017"
            disciplina_auto = "LINGUAGEM E MOVIMENTO"
            
            for l in linhas[:5]:
                if l and "CIEP" in l.upper() and "ANDRE" in l.upper():
                    partes = [p.replace('"', '').strip() for p in re.split(r'[;,]', l) if p.strip()]
                    if len(partes) >= 3:
                        escola_auto = partes[0]
                        disciplina_auto = partes[-1]

            nova_turma = Turma(escola=escola_auto, nome_turma=turma_auto, disciplina=disciplina_auto)
            db.session.add(nova_turma)
            db.session.commit()

            contador_chamada = 1

            for linha in linhas:
                if not linha or not isinstance(linha, str):
                    continue
                    
                linha_up = linha.upper()
                
                if "NUM_CHAMADA" in linha_up or "ANDRE CAMARGO" in linha_up or "CANCELADO" in linha_up:
                    continue
                
                if "MATRICULADO" in linha_up:
                    linha_limpa = linha.replace('"', '')
                    partes = [p.strip() for p in re.split(r'[;,]', linha_limpa) if p.strip()]
                    
                    matricula = ""
                    nome = ""
                    
                    for p in partes:
                        if p.isdigit() and len(p) >= 10:
                            matricula = p
                        elif len(p) > 8 and "MATRICULADO" not in p.upper() and "TRIMESTRE" not in p.upper():
                            nome = p.upper()
                    
                    if nome and matricula:
                        db.session.add(Aluno(
                            num_chamada=str(contador_chamada),
                            matricula=matricula,
                            nome=nome,
                            situacao="MATRICULADO",
                            turma_id=nova_turma.id
                        ))
                        contador_chamada += 1
                        
            db.session.commit()
        except Exception as e:
            print(f"Erro na varredura defensiva: {e}")
            
    return redirect(url_for('index'))

@app.route('/chamada/<int:turma_id>')
def chamada(turma_id):
    turma = Turma.query.get_or_404(turma_id)
    data_filtro = datetime.today().strftime('%Y-%m-%d')
    alunos = Aluno.query.filter_by(turma_id=turma.id).all()
    
    # Ordenação robusta contra valores em branco ou strings
    alunos_ordenados = sorted(alunos, key=lambda x: int(x.num_chamada) if (x.num_chamada and str(x.num_chamada).isdigit()) else 99)
    
    alunos_info = []
    for a in alunos_ordenados:
        reg = Presenca.query.filter_by(aluno_id=a.id, data=data_filtro).first()
        status_hoje = reg.status if reg else 'P'
        alunos_info.append({
            "id": a.id, 
            "num_chamada": a.num_chamada, 
            "matricula": a.matricula, 
            "nome": a.nome, 
            "situacao": a.situacao, 
            "status_hoje": status_hoje
        })
        
    return render_template_string(HTML_COMPLETO, tela='chamada', turma=turma, alunos_info=alunos_info, data_atual=data_filtro)

@app.route('/salvar-chamada/<int:turma_id>', methods=['POST'])
def salvar_chamada(turma_id):
    data_chamada = request.form.get('data_chamada')
    alunos = Aluno.query.filter_by(turma_id=turma_id).all()
    
    for a in alunos:
        status = request.form.get(f'status_{a.id}', 'P')
        # Correção aqui: garantindo a busca correta por ID numérico do aluno
        reg = Presenca.query.filter_by(aluno_id=int(a.id), data=str(data_chamada)).first()
        if reg: 
            reg.status = status
        else: 
            db.session.add(Presenca(data=str(data_chamada), status=status, aluno_id=int(a.id)))
            
    db.session.commit()
    return redirect(url_for('chamada', turma_id=turma_id))

@app.route('/excluir-turma/<int:turma_id>')
def excluir_turma(turma_id):
    turma = Turma.query.get(turma_id)
    if r:= turma: 
        db.session.delete(r)
        db.session.commit()
    return redirect(url_for('index'))

@app.route('/baixar-excel/<int:turma_id>')
def baixar_excel(turma_id):
    try:
        turma = Turma.query.get_or_404(turma_id)
        alunos = Aluno.query.filter_by(turma_id=turma.id).all()
        alunos_ordenados = sorted(alunos, key=lambda x: int(x.num_chamada) if (x.num_chamada and str(x.num_chamada).isdigit()) else 99)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "DIARIO"
        ws.views.sheetView[0].showGridLines = True
        
        headers = ["Nº Chamada", "Matrícula", "Nome Completo do Aluno", "Situação"]
        for col, h in enumerate(headers, 1): 
            ws.cell(row=1, column=col, value=h)
        
        for idx, a in enumerate(alunos_ordenados, 2):
            num = int(a.num_chamada) if (a.num_chamada and str(a.num_chamada).isdigit()) else idx-1
            ws.cell(row=idx, column=1, value=num)
            ws.cell(row=idx, column=2, value=str(a.matricula))
            ws.cell(row=idx, column=3, value=str(a.nome))
            ws.cell(row=idx, column=4, value=str(a.situacao))
            
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(
            output, 
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            as_attachment=True, 
            download_name=f"Diario_Limpo_Turma_{turma.nome_turma}.xlsx"
        )
    except Exception as e:
        return f"Erro ao gerar Excel: {str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

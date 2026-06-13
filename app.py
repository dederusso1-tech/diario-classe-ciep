import os
import re
from datetime import datetime
import pandas as pd
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Configura o banco de dados local estável
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(BASE_DIR, 'diario_escola.db')
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
        <span class="navbar-brand mb-0 h1 text-white">🍎 Diário de Frequência - CIEP 205</span>
        <a href="/" class="btn btn-sm btn-outline-light fw-bold">🏠 Voltar ao Início</a>
    </div>
</nav>

<div class="container">
    {% if tela == 'inicial' %}
    <div class="row">
        <div class="col-md-5 mb-4">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-primary mb-3">📂 Cadastrar Nova Turma (CSV)</h5>
                <form action="/carregar-csv" method="POST" enctype="multipart/form-data">
                    <div class="mb-2">
                        <label class="form-label small fw-bold">Escola</label>
                        <input type="text" name="escola" class="form-control form-control-sm" value="CIEP 205 FREI AGOSTINHO FÍNCIAS" required>
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold">Turma</label>
                        <input type="text" name="turma" class="form-control form-control-sm" placeholder="Ex: 1017" required>
                    </div>
                    <div class="mb-2">
                        <label class="form-label small fw-bold">Disciplina</label>
                        <input type="text" name="disciplina" class="form-control form-control-sm" placeholder="Ex: REDAÇÃO" required>
                    </div>
                    <div class="mb-3">
                        <label class="form-label small fw-bold">Selecione o Arquivo CSV da Escola</label>
                        <input type="file" name="arquivo_csv" class="form-control form-control-sm" accept=".csv" required>
                    </div>
                    <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold">🔄 Limpar CSV e Salvar Turma</button>
                </form>
            </div>
        </div>

        <div class="col-md-7">
            <div class="card card-custom p-4 bg-white">
                <h5 class="fw-bold text-success mb-3">📚 Suas Turmas Ativas</h5>
                {% if not turmas %}
                    <p class="text-muted small">Nenhuma turma carregada nesta sessão. Importe o arquivo CSV ao lado para começar.</p>
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
                <h4 class="fw-bold m-0 text-dark">📋 Diário de Frequência Diária</h4>
                <small class="text-muted">Turma: {{ turma.nome_turma }} | Matéria: {{ turma.disciplina }}</small>
            </div>
            <a href="/baixar-excel/{{ turma.id }}" class="btn btn-success fw-bold px-4 shadow-sm">
                📥 Exportar Planilha Excel Limpa
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
                <table class="table table-striped table-bordered align-middle m-0">
                    <thead class="table-dark">
                        <tr>
                            <th>Nome Completo do Aluno (Filtrado)</th>
                            <th class="text-center" style="width: 280px;">Presença / Falta</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for aluno in alunos_info %}
                        <tr>
                            <td class="ps-3"><strong>{{ aluno.nome }}</strong></td>
                            <td class="text-center">
                                <div class="btn-group" role="group">
                                    <input type="radio" class="btn-check" name="status_{{ aluno.id }}" id="p_{{ aluno.id }}" value="P" {% if aluno.status_hoje == 'P' %}checked{% endif %}>
                                    <label class="btn btn-sm btn-outline-success px-3 fw-bold" for="p_{{ aluno.id }}">P (Presença)</label>

                                    <input type="radio" class="btn-check" name="status_{{ aluno.id }}" id="f_{{ aluno.id }}" value="F" {% if aluno.status_hoje == 'F' %}checked{% endif %}>
                                    <label class="btn btn-sm btn-outline-danger px-3 fw-bold" for="f_{{ aluno.id }}">F (Falta)</label>
                                </div>
                            </td>
                        </tr>
                        {% endfor %}
                    </tbody>
                </table>
            </div>
            <div class="text-end mt-3">
                <button type="submit" class="btn btn-primary fw-bold px-4">💾 Gravar Chamada</button>
            </div>
        </form>
    </div>
    {% endif %}
</div>

</body>
</html>
'''

def limpar_nome_aluno(texto_bruto):
    """ Remove matrículas, códigos de turmas e símbolos, deixando apenas o nome """
    texto = str(texto_bruto).upper().strip()
    # Remove sequências longas de números (como matrículas)
    texto = re.sub(re.compile(r'\b\d{7,}\b'), '', texto)
    # Remove códigos comuns de turmas misturados
    texto = re.sub(re.compile(r'\b\d{4}\b'), '', texto)
    # Remove caracteres e pontuações perdidas
    texto = re.sub(re.compile(r'[;\-,.\/|]'), ' ', texto)
    # Limpa espaços duplos
    texto = " ".join(texto.split())
    return texto

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
            # Lê o arquivo independente do separador nativo do sistema escolar
            try:
                df = pd.read_csv(file, sep=None, engine='python', encoding='utf-8')
            except:
                file.seek(0)
                df = pd.read_csv(file, sep=None, engine='python', encoding='iso-8859-1')

            coluna_nome = df.columns[0]
            for col in df.columns:
                if 'nome' in str(col).lower() or 'aluno' in str(col).lower():
                    coluna_nome = col
                    break

            lista_bruta = df[coluna_nome].dropna().tolist()

            # Cria a estrutura da turma
            nova_turma = Turma(escola=escola, nome_turma=nome_turma, disciplina=disciplina)
            db.session.add(nova_turma)
            db.session.commit()

            for item in lista_bruta:
                nome_limpo = limpar_nome_aluno(item)
                if nome_limpo and not nome_limpo.isnumeric():
                    aluno = Aluno(nome=nome_limpo, turma_id=nova_turma.id)
                    db.session.add(aluno)
            db.session.commit()
        except Exception as e:
            print(f"Erro no processamento: {e}")

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
        alunos_info.append({"id": a.id, "nome": a.nome, "status_hoje": status_hoje})

    return render_template_string(HTML_COMPLETO, tela='chamada', turma=turma, alunos_info=alunos_info, data_atual=data_filtro)

@app.route('/salvar-chamada/<int:turma_id>', methods=['POST'])
def salvar_chamada(turma_id):
    data_chamada = request.form.get('data_chamada')
    turma = Turma.query.get_or_404(turma_id)
    alunos = Aluno.query.filter_by(turma_id=turma.id).all()

    for a in alunos:
        status = request.form.get(f'status_{a.id}', 'P')
        reg = Presenca.query.filter_by(aluno_id=a.id, data=data_chamada).first()
        if reg:
            reg.status = status
        else:
            novo_reg = Presenca(data=data_chamada, status=status, aluno_id=a.id)
            db.session.add(novo_reg)
            
    db.session.commit()
    return redirect(url_for('chamada', turma_id=turma_id, data_filtro=data_chamada))

@app.route('/excluir-turma/<int:turma_id>')
def excluir_turma(turma_id):
    turma = Turma.query.get(turma_id)
    if turma:
        db.session.delete(turma)
        db.session.commit()
    return redirect(url_for('index'))

@app.route('/baixar-excel/<int:turma_id>')
def baixar_excel(turma_id):
    turma = Turma.query.get_or_404(turma_id)
    alunos = Aluno.query.filter_by(turma_id=turma.id).order_by(Aluno.nome).all()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"TURMA {turma.nome_turma}"
    ws.views.sheetView[0].showGridLines = True

    cor_topo = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
    cor_header = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid")
    cor_zebra = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid")
    
    fonte_titulo = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    fonte_header = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    fonte_dados = Font(name="Arial", size=10)
    
    alinhamento_centro = Alignment(horizontal="center", vertical="center")
    alinhamento_esquerda = Alignment(horizontal="left", vertical="center")
    
    borda = Border(left=Side(style="thin", color="CBD5E0"), right=Side(style="thin", color="CBD5E0"),
                   top=Side(style="thin", color="CBD5E0"), bottom=Side(style="thin", color="CBD5E0"))

    # Cabeçalho Fixo Estruturado
    ws.merge_cells("A1:G1")
    ws["A1"] = f"{turma.escola} | DIÁRIO OFICIAL DE CLASSE"
    ws["A1"].font = fonte_titulo
    ws["A1"].fill = cor_topo
    ws["A1"].alignment = alinhamento_centro
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:G2")
    ws["A2"] = f"TURMA: {turma.nome_turma}  |  DISCIPLINA: {turma.disciplina}  |  PROFESSOR: ANDRÉ CAMARGO"
    ws["A2"].font = Font(name="Arial", size=9, italic=True, color="FFFFFF")
    ws["A2"].fill = cor_topo
    ws["A2"].alignment = alinhamento_centro
    ws.row_dimensions[2].height = 18

    headers = ["Nome Completo do Aluno", "Dias Letivos", "Faltas Acumuladas", "Frequência (%)", "Nota 1", "Nota 2", "Média Final"]
    for col_num, text in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num, value=text)
        cell.font = fonte_header
        cell.fill = cor_header
        cell.alignment = alinhamento_centro
        cell.border = borda
    ws.row_dimensions[4].height = 22

    for idx, a in enumerate(alunos):
        r = 5 + idx
        fill_linha = cor_zebra if idx % 2 == 0 else PatternFill(fill_type=None)

        datas_gravadas = db.session.query(Presenca.data).join(Aluno).filter(Aluno.turma_id == turma.id).distinct().count()
        total_dias = max(datas_gravadas, 1)
        total_faltas = Presenca.query.filter_by(aluno_id=a.id, status='F').count()
        porcentagem_freq = ((total_dias - total_faltas) / total_dias)

        c_nome = ws.cell(row=r, column=1, value=a.nome)
        c_dias = ws.cell(row=r, column=2, value=total_dias)
        c_faltas = ws.cell(row=r, column=3, value=total_faltas)
        c_freq = ws.cell(row=r, column=4, value=porcentagem_freq)
        c_freq.number_format = '0%'

        ws.cell(row=r, column=5, value="")
        ws.cell(row=r, column=6, value="")
        ws.cell(row=r, column=7, value="")

        c_nome.font = fonte_dados
        c_nome.alignment = alinhamento_esquerda
        
        for c_idx in range(1, 8):
            cell = ws.cell(row=r, column=c_idx)
            cell.fill = fill_linha
            cell.border = borda
            if c_idx > 1:
                cell.alignment = alinhamento_centro
                cell.font = fonte_dados
        ws.row_dimensions[r].height = 18

    ws.column_dimensions['A'].width = 45
    for col in ['B', 'C', 'D', 'E', 'F', 'G']:
        ws.column_dimensions[col].width = 16

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f"Diario_Limpo_Turma_{turma.nome_turma}.xlsx")

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

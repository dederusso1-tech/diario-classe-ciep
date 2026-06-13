import os
from datetime import datetime
import pandas as pd
from io import BytesIO
from flask import Flask, render_template_string, request, redirect, url_for, send_file
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)

dados_sistema = {
    "escola": "CIEP 205 FREI AGOSTINHO FÍNCIAS",
    "turma": "1017",
    "disciplina": "REDAÇÃO",
    "alunos": [],       
    "frequencia": {}    
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
                📥 Passo 3: Baixar Planilha Excel Formatada
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
    frequencia_hoje = dados_sistema["frequencia"].get(data_filtro, {})
    return render_template_string(HTML_COMPLETO, info=dados_sistema, data_atual=data_filtro, frequencia_hoje=frequencia_hoje, enumerate=enumerate)

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
            dados_sistema["frequencia"] = {} 
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
    # Inicializa o workbook do openpyxl para controle total do design
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"TURMA {dados_sistema['turma']}"
    ws.views.sheetView[0].showGridLines = True

    # Cores e Estilos
    cor_topo = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid") # Azul Marinho
    cor_header = PatternFill(start_color="2B6CB0", end_color="2B6CB0", fill_type="solid") # Azul Médio
    cor_zebra = PatternFill(start_color="F7FAFC", end_color="F7FAFC", fill_type="solid") # Cinza claro
    cor_branco = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    fonte_titulo = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    fonte_sub = Font(name="Arial", size=10, italic=True, color="FFFFFF")
    fonte_header = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    fonte_dados = Font(name="Arial", size=10)
    fonte_dados_bold = Font(name="Arial", size=10, bold=True)
    
    alinhamento_centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alinhamento_esquerda = Alignment(horizontal="left", vertical="center")
    
    borda_fina = Side(border_style="thin", color="CBD5E0")
    border_total = Border(left=borda_fina, right=borda_fina, top=borda_fina, bottom=borda_fina)

    # 1. Montagem do Cabeçalho Oficial Escolar
    ws.merge_cells("A1:G1")
    ws["A1"] = dados_sistema["escola"]
    ws["A1"].font = fonte_titulo
    ws["A1"].fill = cor_topo
    ws["A1"].alignment = alinhamento_centro
    ws.row_dimensions[1].height = 30

    ws.merge_cells("A2:G2")
    ws["A2"] = f"DIÁRIO DE CLASSE - DISCIPLINA: {dados_sistema['disciplina']} | TURMA: {dados_sistema['turma']} | PROFESSOR: ANDRÉ CAMARGO"
    ws["A2"].font = fonte_sub
    ws["A2"].fill = cor_topo
    ws["A2"].alignment = alinhamento_centro
    ws.row_dimensions[2].height = 20

    # Linha em branco de respiro
    ws.row_dimensions[3].height = 15

    # 2. Cabeçalho da Tabela de Dados
    headers = ["Nome Completo do Aluno", "Dias Letivos", "Faltas", "Freq. (%)", "Nota 1", "Nota 2", "Média Final"]
    for col_num, header_text in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num, value=header_text)
        cell.font = fonte_header
        cell.fill = cor_header
        cell.alignment = alinhamento_centro
        cell.border = border_total
    ws.row_dimensions[4].height = 25

    # 3. Inserção dos Alunos e Regras de Negócio
    total_dias_letivos = max(len(dados_sistema["frequencia"]), 1)
    row_start = 5

    for idx, nome in enumerate(dados_sistema["alunos"]):
        current_row = row_start + idx
        fill_atual = cor_zebra if idx % 2 == 0 else cor_branco
        
        # Conta faltas salvas
        faltas_aluno = 0
        for data, chamadas in dados_sistema["frequencia"].items():
            if chamadas.get(str(idx)) == 'F':
                faltas_aluno += 1

        # Escreve os valores nas células
        c_nome = ws.cell(row=current_row, column=1, value=nome)
        c_dias = ws.cell(row=current_row, column=2, value=total_dias_letivos)
        c_faltas = ws.cell(row=current_row, column=3, value=faltas_aluno)
        
        # Fórmula da Frequência: =((Dias Letivos - Faltas) / Dias Letivos)
        c_freq = ws.cell(row=current_row, column=4, value=f"=({total_dias_letivos}-C{current_row})/{total_dias_letivos}")
        c_freq.number_format = '0.0%'
        
        # Campos vazios para o professor preencher notas futuramente
        c_n1 = ws.cell(row=current_row, column=5, value="")
        c_n2 = ws.cell(row=current_row, column=6, value="")
        
        # Fórmula da Média Final: =MÉDIA(Nota1; Nota2)
        c_media = ws.cell(row=current_row, column=7, value=f"=AVERAGE(E{current_row}:F{current_row})")
        c_media.number_format = '0.0'

        # Formatação de Estilos das Linhas
        c_nome.alignment = alinhamento_esquerda
        c_nome.font = fonte_dados_bold
        
        for col_idx in range(1, 8):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.fill = fill_atual
            cell.border = border_total
            if col_idx > 1:
                cell.alignment = alinhamento_centro
                cell.font = fonte_dados
        
        ws.row_dimensions[current_row].height = 20

    # 4. Ajuste Automático de Largura das Colunas para não cortar os nomes
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row > 2 and cell.value: # Ignora o título mesclado para o cálculo do tamanho
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
    
    # Ajuste fixo especial para a coluna de Nomes
    ws.column_dimensions['A'].width = 45

    # Envio do arquivo formatado
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    nome_arquivo = f"Diario_Formatado_{dados_sistema['turma']}.xlsx"
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=nome_arquivo
    )

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

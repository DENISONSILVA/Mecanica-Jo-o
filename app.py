"""
EstoqueOficina — Backend Flask + SQLite
Funcionalidades: Peças, Fornecedores, Movimentações, Ordens de Serviço,
                 Exportação Excel e PDF
"""

import sqlite3, os, io, json
from datetime import datetime
from flask import Flask, request, jsonify, g, send_file, render_template

# ─── Dependências opcionais para exportação ───────────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    PDF_OK = True
except ImportError:
    PDF_OK = False

# ─── App ──────────────────────────────────────────────────────────────────────
app = Flask(__name__)
DB_PATH = os.path.join(os.path.dirname(__file__), "oficina.db")

# ─── DB helpers ───────────────────────────────────────────────────────────────
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db:
        db.close()

def query(sql, params=(), one=False):
    cur = get_db().execute(sql, params)
    r = cur.fetchone() if one else cur.fetchall()
    return (dict(r) if r else None) if one else [dict(row) for row in r]

def execute(sql, params=()):
    db = get_db()
    cur = db.execute(sql, params)
    db.commit()
    return cur.lastrowid

# ─── Inicialização do banco ────────────────────────────────────────────────────
def init_db():
    db = sqlite3.connect(DB_PATH)
    db.execute("PRAGMA foreign_keys = ON")
    db.executescript("""
    CREATE TABLE IF NOT EXISTS fornecedores (
        id        INTEGER PRIMARY KEY AUTOINCREMENT,
        nome      TEXT NOT NULL,
        cnpj      TEXT,
        telefone  TEXT,
        email     TEXT,
        contato   TEXT,
        endereco  TEXT,
        obs       TEXT,
        criado_em TEXT DEFAULT (datetime('now','localtime'))
    );

    CREATE TABLE IF NOT EXISTS pecas (
        id           INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo       TEXT NOT NULL UNIQUE,
        nome         TEXT NOT NULL,
        categoria    TEXT,
        qtd          INTEGER DEFAULT 0,
        qtd_min      INTEGER DEFAULT 0,
        preco        REAL DEFAULT 0,
        local        TEXT,
        fornecedor_id INTEGER REFERENCES fornecedores(id) ON DELETE SET NULL,
        criado_em    TEXT DEFAULT (datetime('now','localtime'))
    );

    CREATE TABLE IF NOT EXISTS movimentacoes (
        id        INTEGER PRIMARY KEY AUTOINCREMENT,
        peca_id   INTEGER NOT NULL REFERENCES pecas(id),
        os_id     INTEGER REFERENCES ordens_servico(id) ON DELETE SET NULL,
        tipo      TEXT NOT NULL CHECK(tipo IN ('entrada','saida')),
        qtd       INTEGER NOT NULL,
        obs       TEXT,
        data      TEXT DEFAULT (datetime('now','localtime'))
    );

    CREATE TABLE IF NOT EXISTS ordens_servico (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        numero      TEXT NOT NULL UNIQUE,
        cliente     TEXT NOT NULL,
        veiculo     TEXT,
        placa       TEXT,
        servico     TEXT,
        status      TEXT DEFAULT 'aberta' CHECK(status IN ('aberta','em_andamento','concluida','cancelada')),
        valor_mao   REAL DEFAULT 0,
        obs         TEXT,
        criado_em   TEXT DEFAULT (datetime('now','localtime')),
        fechado_em  TEXT
    );

    CREATE TABLE IF NOT EXISTS os_pecas (
        id      INTEGER PRIMARY KEY AUTOINCREMENT,
        os_id   INTEGER NOT NULL REFERENCES ordens_servico(id) ON DELETE CASCADE,
        peca_id INTEGER NOT NULL REFERENCES pecas(id),
        qtd     INTEGER NOT NULL,
        preco   REAL NOT NULL
    );
    """)
    db.commit()

    # Seed se vazio
    count = db.execute("SELECT COUNT(*) FROM pecas").fetchone()[0]
    if count == 0:
        db.executescript("""
        INSERT INTO fornecedores(nome,cnpj,telefone,email,contato) VALUES
          ('AutoParts BA','12.345.678/0001-90','(71) 3333-1111','vendas@autopartsba.com','João Silva'),
          ('Meca Sul','98.765.432/0001-10','(71) 3333-2222','contato@mecasul.com','Maria Souza'),
          ('Distribuidora X','11.222.333/0001-44','(71) 3333-3333','dx@distribuidorax.com','Pedro Lima'),
          ('BatMax','44.555.666/0001-77','(71) 3333-4444','batmax@batmax.com','Ana Costa'),
          ('NGK Brasil','77.888.999/0001-11','(71) 3333-5555','ngk@ngk.com','Carlos Reis');

        INSERT INTO pecas(codigo,nome,categoria,qtd,qtd_min,preco,local,fornecedor_id) VALUES
          ('FRE-001','Pastilha de Freio Dianteira','Freios',8,4,45.90,'A1',1),
          ('FRE-002','Disco de Freio Traseiro','Freios',2,2,189.00,'A2',1),
          ('MOT-001','Correia Dentada','Motor',3,2,78.50,'B1',2),
          ('FIL-001','Filtro de Óleo','Filtros',1,5,18.90,'C1',3),
          ('FIL-002','Filtro de Ar','Filtros',6,3,24.00,'C2',3),
          ('ELE-001','Bateria 60Ah','Elétrica',0,1,320.00,'D1',4),
          ('SUS-001','Amortecedor Dianteiro','Suspensão',4,2,215.00,'E1',2),
          ('MOT-002','Vela de Ignição (un.)','Motor',16,8,22.00,'B2',5);

        INSERT INTO ordens_servico(numero,cliente,veiculo,placa,servico,status,valor_mao) VALUES
          ('OS-0001','Carlos Mendes','Gol 2019','BBA-1234','Troca de pastilhas e revisão de freios','concluida',180.00),
          ('OS-0002','Fernanda Lima','HB20 2021','CCB-5678','Troca de correia dentada','em_andamento',250.00),
          ('OS-0003','Roberto Alves','Onix 2020','DDD-9012','Revisão geral','aberta',0.00);

        INSERT INTO os_pecas(os_id,peca_id,qtd,preco) VALUES
          (1,1,4,45.90),(1,2,1,189.00),
          (2,3,1,78.50),
          (3,4,2,18.90),(3,5,1,24.00);

        INSERT INTO movimentacoes(peca_id,os_id,tipo,qtd,obs) VALUES
          (1,1,'saida',4,'OS-0001 — pastilhas'),
          (2,1,'saida',1,'OS-0001 — disco'),
          (3,NULL,'entrada',3,'Compra fornecedor Meca Sul'),
          (4,NULL,'saida',3,'OS-0003 — filtros'),
          (5,NULL,'entrada',6,'Reposição estoque');
        """)
        db.commit()
    db.close()

# ═══════════════════════════════════════════════════════
#  FRONTEND
# ═══════════════════════════════════════════════════════
@app.route("/")
def index():
    return render_template("index.html")

# ═══════════════════════════════════════════════════════
#  PEÇAS
# ═══════════════════════════════════════════════════════
@app.route("/api/pecas", methods=["GET"])
def get_pecas():
    rows = query("""
        SELECT p.*, f.nome AS fornecedor_nome
        FROM pecas p
        LEFT JOIN fornecedores f ON p.fornecedor_id = f.id
        ORDER BY p.nome
    """)
    return jsonify(rows)

@app.route("/api/pecas", methods=["POST"])
def create_peca():
    d = request.json
    pid = execute("""
        INSERT INTO pecas(codigo,nome,categoria,qtd,qtd_min,preco,local,fornecedor_id)
        VALUES(?,?,?,?,?,?,?,?)
    """, (d["codigo"], d["nome"], d.get("categoria",""), int(d.get("qtd",0)),
          int(d.get("qtd_min",0)), float(d.get("preco",0)),
          d.get("local",""), d.get("fornecedor_id") or None))
    return jsonify({"id": pid}), 201

@app.route("/api/pecas/<int:pid>", methods=["PUT"])
def update_peca(pid):
    d = request.json
    execute("""
        UPDATE pecas SET codigo=?,nome=?,categoria=?,qtd=?,qtd_min=?,preco=?,local=?,fornecedor_id=?
        WHERE id=?
    """, (d["codigo"], d["nome"], d.get("categoria",""), int(d.get("qtd",0)),
          int(d.get("qtd_min",0)), float(d.get("preco",0)),
          d.get("local",""), d.get("fornecedor_id") or None, pid))
    return jsonify({"ok": True})

@app.route("/api/pecas/<int:pid>", methods=["DELETE"])
def delete_peca(pid):
    execute("DELETE FROM pecas WHERE id=?", (pid,))
    return jsonify({"ok": True})

# ═══════════════════════════════════════════════════════
#  MOVIMENTAÇÕES
# ═══════════════════════════════════════════════════════
@app.route("/api/movimentacoes", methods=["GET"])
def get_movs():
    rows = query("""
        SELECT m.*, p.nome AS peca_nome, p.codigo AS peca_codigo,
               o.numero AS os_numero
        FROM movimentacoes m
        JOIN pecas p ON m.peca_id = p.id
        LEFT JOIN ordens_servico o ON m.os_id = o.id
        ORDER BY m.data DESC LIMIT 100
    """)
    return jsonify(rows)

@app.route("/api/movimentacoes", methods=["POST"])
def create_mov():
    d = request.json
    peca = query("SELECT * FROM pecas WHERE id=?", (d["peca_id"],), one=True)
    if not peca:
        return jsonify({"erro": "Peça não encontrada"}), 404
    qtd = int(d["qtd"])
    if d["tipo"] == "saida" and qtd > peca["qtd"]:
        return jsonify({"erro": f"Estoque insuficiente. Disponível: {peca['qtd']} un."}), 400

    nova = peca["qtd"] + qtd if d["tipo"] == "entrada" else peca["qtd"] - qtd
    execute("UPDATE pecas SET qtd=? WHERE id=?", (nova, d["peca_id"]))
    mid = execute("""
        INSERT INTO movimentacoes(peca_id,os_id,tipo,qtd,obs)
        VALUES(?,?,?,?,?)
    """, (d["peca_id"], d.get("os_id") or None, d["tipo"], qtd, d.get("obs","")))
    return jsonify({"id": mid, "qtd_atual": nova}), 201

# ═══════════════════════════════════════════════════════
#  FORNECEDORES
# ═══════════════════════════════════════════════════════
@app.route("/api/fornecedores", methods=["GET"])
def get_fornecedores():
    rows = query("""
        SELECT f.*, COUNT(p.id) AS total_pecas
        FROM fornecedores f
        LEFT JOIN pecas p ON p.fornecedor_id = f.id
        GROUP BY f.id ORDER BY f.nome
    """)
    return jsonify(rows)

@app.route("/api/fornecedores", methods=["POST"])
def create_fornecedor():
    d = request.json
    fid = execute("""
        INSERT INTO fornecedores(nome,cnpj,telefone,email,contato,endereco,obs)
        VALUES(?,?,?,?,?,?,?)
    """, (d["nome"], d.get("cnpj",""), d.get("telefone",""),
          d.get("email",""), d.get("contato",""), d.get("endereco",""), d.get("obs","")))
    return jsonify({"id": fid}), 201

@app.route("/api/fornecedores/<int:fid>", methods=["PUT"])
def update_fornecedor(fid):
    d = request.json
    execute("""
        UPDATE fornecedores SET nome=?,cnpj=?,telefone=?,email=?,contato=?,endereco=?,obs=?
        WHERE id=?
    """, (d["nome"], d.get("cnpj",""), d.get("telefone",""),
          d.get("email",""), d.get("contato",""), d.get("endereco",""), d.get("obs",""), fid))
    return jsonify({"ok": True})

@app.route("/api/fornecedores/<int:fid>", methods=["DELETE"])
def delete_fornecedor(fid):
    execute("DELETE FROM fornecedores WHERE id=?", (fid,))
    return jsonify({"ok": True})

# ═══════════════════════════════════════════════════════
#  ORDENS DE SERVIÇO
# ═══════════════════════════════════════════════════════
@app.route("/api/os", methods=["GET"])
def get_os():
    rows = query("""
        SELECT o.*,
               COALESCE(SUM(op.qtd * op.preco),0) AS valor_pecas,
               COUNT(op.id) AS total_itens
        FROM ordens_servico o
        LEFT JOIN os_pecas op ON op.os_id = o.id
        GROUP BY o.id ORDER BY o.criado_em DESC
    """)
    return jsonify(rows)

@app.route("/api/os/<int:oid>", methods=["GET"])
def get_os_detail(oid):
    os = query("SELECT * FROM ordens_servico WHERE id=?", (oid,), one=True)
    if not os:
        return jsonify({"erro": "OS não encontrada"}), 404
    pecas = query("""
        SELECT op.*, p.nome AS peca_nome, p.codigo AS peca_codigo
        FROM os_pecas op JOIN pecas p ON op.peca_id = p.id
        WHERE op.os_id=?
    """, (oid,))
    os["pecas"] = pecas
    return jsonify(os)

@app.route("/api/os", methods=["POST"])
def create_os():
    d = request.json
    # Gera número automático
    ultimo = query("SELECT numero FROM ordens_servico ORDER BY id DESC LIMIT 1", one=True)
    if ultimo:
        try:
            n = int(ultimo["numero"].split("-")[-1]) + 1
        except:
            n = 1
    else:
        n = 1
    numero = f"OS-{str(n).zfill(4)}"
    oid = execute("""
        INSERT INTO ordens_servico(numero,cliente,veiculo,placa,servico,status,valor_mao,obs)
        VALUES(?,?,?,?,?,?,?,?)
    """, (numero, d["cliente"], d.get("veiculo",""), d.get("placa",""),
          d.get("servico",""), d.get("status","aberta"),
          float(d.get("valor_mao",0)), d.get("obs","")))
    return jsonify({"id": oid, "numero": numero}), 201

@app.route("/api/os/<int:oid>", methods=["PUT"])
def update_os(oid):
    d = request.json
    fechado = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if d.get("status") == "concluida" else None
    execute("""
        UPDATE ordens_servico
        SET cliente=?,veiculo=?,placa=?,servico=?,status=?,valor_mao=?,obs=?,
            fechado_em=COALESCE(fechado_em,?)
        WHERE id=?
    """, (d["cliente"], d.get("veiculo",""), d.get("placa",""), d.get("servico",""),
          d.get("status","aberta"), float(d.get("valor_mao",0)), d.get("obs",""),
          fechado, oid))
    return jsonify({"ok": True})

@app.route("/api/os/<int:oid>", methods=["DELETE"])
def delete_os(oid):
    execute("DELETE FROM ordens_servico WHERE id=?", (oid,))
    return jsonify({"ok": True})

@app.route("/api/os/<int:oid>/pecas", methods=["POST"])
def add_peca_os(oid):
    d = request.json
    peca = query("SELECT * FROM pecas WHERE id=?", (d["peca_id"],), one=True)
    if not peca:
        return jsonify({"erro": "Peça não encontrada"}), 404
    qtd = int(d["qtd"])
    if qtd > peca["qtd"]:
        return jsonify({"erro": f"Estoque insuficiente. Disponível: {peca['qtd']} un."}), 400

    execute("INSERT INTO os_pecas(os_id,peca_id,qtd,preco) VALUES(?,?,?,?)",
            (oid, d["peca_id"], qtd, float(d.get("preco", peca["preco"]))))
    # Dá baixa automática
    execute("UPDATE pecas SET qtd=qtd-? WHERE id=?", (qtd, d["peca_id"]))
    execute("""INSERT INTO movimentacoes(peca_id,os_id,tipo,qtd,obs)
               VALUES(?,?,?,?,?)""",
            (d["peca_id"], oid, "saida", qtd, f"Vinculado à OS #{oid}"))
    return jsonify({"ok": True}), 201

@app.route("/api/os/<int:oid>/pecas/<int:opid>", methods=["DELETE"])
def remove_peca_os(oid, opid):
    item = query("SELECT * FROM os_pecas WHERE id=? AND os_id=?", (opid, oid), one=True)
    if not item:
        return jsonify({"erro": "Item não encontrado"}), 404
    # Estorna estoque
    execute("UPDATE pecas SET qtd=qtd+? WHERE id=?", (item["qtd"], item["peca_id"]))
    execute("DELETE FROM os_pecas WHERE id=?", (opid,))
    return jsonify({"ok": True})

# ═══════════════════════════════════════════════════════
#  EXPORTAÇÃO — EXCEL
# ═══════════════════════════════════════════════════════
@app.route("/api/export/excel")
def export_excel():
    if not EXCEL_OK:
        return jsonify({"erro": "openpyxl não instalado. Rode: pip install openpyxl"}), 500

    wb = openpyxl.Workbook()

    # ─ Aba Estoque ─
    ws = wb.active
    ws.title = "Estoque"
    header_fill = PatternFill("solid", fgColor="1E3A5F")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    cols = ["Código","Nome","Categoria","Qtd","Mín.","Status","Preço (R$)","Localização","Fornecedor"]
    for ci, col in enumerate(cols, 1):
        c = ws.cell(1, ci, col)
        c.fill = header_fill; c.font = header_font
        c.alignment = Alignment(horizontal="center")
    pecas = query("SELECT p.*, f.nome AS fnome FROM pecas p LEFT JOIN fornecedores f ON p.fornecedor_id=f.id ORDER BY p.nome")
    for ri, p in enumerate(pecas, 2):
        st = "Zerado" if p["qtd"]==0 else ("Baixo" if p["qtd"]<=p["qtd_min"] else "Ok")
        ws.append([p["codigo"],p["nome"],p["categoria"],p["qtd"],p["qtd_min"],st,p["preco"],p["local"],p["fnome"] or ""])
    ws.column_dimensions["B"].width = 36
    ws.column_dimensions["I"].width = 22

    # ─ Aba OS ─
    ws2 = wb.create_sheet("Ordens de Serviço")
    cols2 = ["Número","Cliente","Veículo","Placa","Serviço","Status","Mão de Obra","Peças (R$)","Total (R$)","Aberta em"]
    for ci, col in enumerate(cols2, 1):
        c = ws2.cell(1, ci, col)
        c.fill = header_fill; c.font = header_font
    os_list = query("""
        SELECT o.*, COALESCE(SUM(op.qtd*op.preco),0) AS vp
        FROM ordens_servico o LEFT JOIN os_pecas op ON op.os_id=o.id
        GROUP BY o.id ORDER BY o.criado_em DESC
    """)
    for ri, o in enumerate(os_list, 2):
        total = o["valor_mao"] + o["vp"]
        ws2.append([o["numero"],o["cliente"],o["veiculo"],o["placa"],o["servico"],
                    o["status"],o["valor_mao"],o["vp"],total,o["criado_em"]])
    ws2.column_dimensions["E"].width = 30

    # ─ Aba Movimentações ─
    ws3 = wb.create_sheet("Movimentações")
    cols3 = ["Data","Peça","Código","Tipo","Qtd","OS","Observação"]
    for ci, col in enumerate(cols3, 1):
        c = ws3.cell(1, ci, col)
        c.fill = header_fill; c.font = header_font
    movs = query("""
        SELECT m.*,p.nome AS pn,p.codigo AS pc,o.numero AS on2
        FROM movimentacoes m JOIN pecas p ON m.peca_id=p.id
        LEFT JOIN ordens_servico o ON m.os_id=o.id
        ORDER BY m.data DESC
    """)
    for m in movs:
        ws3.append([m["data"],m["pn"],m["pc"],m["tipo"].upper(),m["qtd"],m["on2"] or "",m["obs"] or ""])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"estoque_oficina_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=fname)

# ═══════════════════════════════════════════════════════
#  EXPORTAÇÃO — PDF
# ═══════════════════════════════════════════════════════
@app.route("/api/export/pdf")
def export_pdf():
    if not PDF_OK:
        return jsonify({"erro": "reportlab não instalado. Rode: pip install reportlab"}), 500

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.5*cm, rightMargin=1.5*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    story  = []

    title_style = ParagraphStyle("title", fontSize=18, fontName="Helvetica-Bold",
                                  spaceAfter=4, alignment=TA_CENTER)
    sub_style   = ParagraphStyle("sub",   fontSize=10, fontName="Helvetica",
                                  textColor=colors.grey, spaceAfter=16, alignment=TA_CENTER)
    h2_style    = ParagraphStyle("h2",    fontSize=13, fontName="Helvetica-Bold", spaceAfter=8, spaceBefore=16)

    story.append(Paragraph("EstoqueOficina", title_style))
    story.append(Paragraph(f"Relatório gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}", sub_style))

    # ─ Resumo ─
    pecas = query("SELECT * FROM pecas")
    total_val  = sum(p["preco"]*p["qtd"] for p in pecas)
    total_low  = sum(1 for p in pecas if 0 < p["qtd"] <= p["qtd_min"])
    total_out  = sum(1 for p in pecas if p["qtd"] == 0)
    story.append(Paragraph("Resumo do Estoque", h2_style))
    resumo_data = [
        ["Total de itens", "Valor em estoque", "Estoque baixo", "Sem estoque"],
        [str(len(pecas)), f"R$ {total_val:,.2f}", str(total_low), str(total_out)],
    ]
    t = Table(resumo_data, colWidths=[4*cm]*4)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#1E3A5F")),
        ("TEXTCOLOR",(0,0),(-1,0), colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("ALIGN",(0,0),(-1,-1),"CENTER"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.HexColor("#F0F4F8"), colors.white]),
        ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#CCCCCC")),
        ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#DDDDDD")),
        ("TOPPADDING",(0,0),(-1,-1),6), ("BOTTOMPADDING",(0,0),(-1,-1),6),
    ]))
    story.append(t)

    # ─ Estoque ─
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("Estoque de Peças", h2_style))
    pecas_full = query("SELECT p.*,f.nome AS fn FROM pecas p LEFT JOIN fornecedores f ON p.fornecedor_id=f.id ORDER BY p.nome")
    pdata = [["Código","Nome","Categoria","Qtd","Mín","Preço","Status"]]
    for p in pecas_full:
        st = "Zerado" if p["qtd"]==0 else ("Baixo" if p["qtd"]<=p["qtd_min"] else "Ok")
        pdata.append([p["codigo"],p["nome"][:30],p["categoria"],str(p["qtd"]),str(p["qtd_min"]),f'R${p["preco"]:.2f}',st])
    t2 = Table(pdata, colWidths=[2.5*cm,6.5*cm,3*cm,1.5*cm,1.5*cm,2.5*cm,1.8*cm])
    style_list = [
        ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#1E3A5F")),
        ("TEXTCOLOR",(0,0),(-1,0), colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.HexColor("#F7F9FC"), colors.white]),
        ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#BBBBBB")),
        ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#DDDDDD")),
        ("TOPPADDING",(0,0),(-1,-1),4), ("BOTTOMPADDING",(0,0),(-1,-1),4),
    ]
    for ri, p in enumerate(pecas_full, 1):
        if p["qtd"] == 0:
            style_list.append(("TEXTCOLOR",(6,ri),(6,ri), colors.red))
        elif p["qtd"] <= p["qtd_min"]:
            style_list.append(("TEXTCOLOR",(6,ri),(6,ri), colors.orange))
        else:
            style_list.append(("TEXTCOLOR",(6,ri),(6,ri), colors.green))
    t2.setStyle(TableStyle(style_list))
    story.append(t2)

    # ─ OS ─
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph("Ordens de Serviço", h2_style))
    os_list = query("""
        SELECT o.*,COALESCE(SUM(op.qtd*op.preco),0) AS vp
        FROM ordens_servico o LEFT JOIN os_pecas op ON op.os_id=o.id
        GROUP BY o.id ORDER BY o.criado_em DESC
    """)
    odata = [["Número","Cliente","Veículo","Status","Mão Obra","Peças","Total"]]
    for o in os_list:
        tot = o["valor_mao"] + o["vp"]
        odata.append([o["numero"],o["cliente"][:20],o["veiculo"],o["status"],
                      f'R${o["valor_mao"]:.2f}',f'R${o["vp"]:.2f}',f'R${tot:.2f}'])
    t3 = Table(odata, colWidths=[2.5*cm,5*cm,3.5*cm,3*cm,2.5*cm,2.5*cm,2.5*cm])
    t3.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#1E3A5F")),
        ("TEXTCOLOR",(0,0),(-1,0), colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.HexColor("#F7F9FC"), colors.white]),
        ("BOX",(0,0),(-1,-1),0.5,colors.HexColor("#BBBBBB")),
        ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#DDDDDD")),
        ("TOPPADDING",(0,0),(-1,-1),4), ("BOTTOMPADDING",(0,0),(-1,-1),4),
    ]))
    story.append(t3)

    doc.build(story)
    buf.seek(0)
    fname = f"estoque_oficina_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    return send_file(buf, mimetype="application/pdf", as_attachment=True, download_name=fname)

# ═══════════════════════════════════════════════════════
#  DASHBOARD STATS
# ═══════════════════════════════════════════════════════
@app.route("/api/stats")
def get_stats():
    pecas  = query("SELECT * FROM pecas")
    os_all = query("SELECT * FROM ordens_servico")
    movs   = query("SELECT * FROM movimentacoes ORDER BY data DESC LIMIT 5")
    os_v   = query("SELECT COALESCE(SUM(op.qtd*op.preco),0)+COALESCE(SUM(o.valor_mao),0) AS total FROM ordens_servico o LEFT JOIN os_pecas op ON op.os_id=o.id", one=True)
    return jsonify({
        "total_pecas":    len(pecas),
        "valor_estoque":  sum(p["preco"]*p["qtd"] for p in pecas),
        "estoque_baixo":  sum(1 for p in pecas if 0 < p["qtd"] <= p["qtd_min"]),
        "estoque_zerado": sum(1 for p in pecas if p["qtd"] == 0),
        "total_os":       len(os_all),
        "os_abertas":     sum(1 for o in os_all if o["status"] in ("aberta","em_andamento")),
        "faturamento_os": os_v["total"] if os_v else 0,
    })

# ═══════════════════════════════════════════════════════
if __name__ == "__main__":
    init_db()
    print("\n✅  EstoqueOficina iniciado!")
    print("🌐  Acesse: http://192.168.1.64:5000\n")
    app.run(debug=True, port=5000)
"""Gera o relatório Word (.docx) a partir dos mesmos motores do dashboard.

Rodar:  python gerar_relatorio.py
Saída:  relatorio/Relatorio_Projeto_Final_PO.docx

Trechos marcados em amarelo com [PREENCHER] ou [CONFERIR] são da equipe.
Ao abrir no Word, clique com o botão direito no Sumário > Atualizar campo.
"""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from motores.ahp import calcular_ahp, matriz_de_triangulo, normalizar_min_max, ranking
from motores.simplex import fmt, resolver

BASE = Path(__file__).parent
SAIDA_DIR = BASE / "relatorio"
GRAF_DIR = SAIDA_DIR / "graficos"
CAMBIO = 5.4701
CRITERIOS = ["Custo mensal (R$)", "Latência SP (ms)", "SLA (%)", "IOPS", "Suporte (1-5)"]
CURTOS = ["Custo", "Latência", "SLA", "IOPS", "Suporte"]
TIPOS = {"Custo mensal (R$)": "custo", "Latência SP (ms)": "custo",
         "SLA (%)": "beneficio", "IOPS": "beneficio", "Suporte (1-5)": "beneficio"}
JULG = {(0, 1): 3, (0, 2): 3, (0, 3): 5, (0, 4): 5, (1, 2): 1, (1, 3): 3, (1, 4): 3,
        (2, 3): 3, (2, 4): 3, (3, 4): 1}

# Etapa 2 (mesmos valores padrão do dashboard)
C = [500, 1200]
A = [[2, 4], [8, 32], [148, 391]]
B = [64, 320, 6000]
RECURSOS = ["vCPU (cota)", "RAM (GB)", "Orçamento (R$/mês)"]
SIGNIF_FOLGA = ["vCPUs da cota não alocados", "GB de RAM não alocados", "R$ do orçamento mensal não gastos"]

# ------------------------------------------------------------------ identidade visual
AZUL = "1F3A5F"        # títulos, cabeçalho de tabela
AZUL_RGB = RGBColor(0x1F, 0x3A, 0x5F)
TEAL = "2A9D8F"        # destaque
TEAL_RGB = RGBColor(0x2A, 0x9D, 0x8F)
CINZA_TXT = RGBColor(0x55, 0x5F, 0x6B)
ZEBRA = "F2F5F9"
BORDA = "C9D3DF"
AMARELO_BOX = "FFF8DC"
AMARELO_BORDA = "E0B100"
PIVO_COL = "FFF1C2"
PIVO_LIN = "DCEBFA"
PIVO_CEL = "F4A259"
FONTE = "Arial"


# ------------------------------------------------------------------ helpers de XML
def sombrear(cel, cor):
    tcPr = cel._tc.get_or_add_tcPr()
    for s in tcPr.findall(qn("w:shd")):
        tcPr.remove(s)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), cor)
    tcPr.append(shd)


def bordas_celula(cel, **lados):
    """lados: top/left/bottom/right = (tamanho_oitavos_pt, cor) ou None para sem borda."""
    tcPr = cel._tc.get_or_add_tcPr()
    tb = tcPr.find(qn("w:tcBorders"))
    if tb is None:
        tb = OxmlElement("w:tcBorders")
        tcPr.append(tb)
    for lado, val in lados.items():
        el = OxmlElement(f"w:{lado}")
        if val is None:
            el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(val[0]))
            el.set(qn("w:color"), val[1])
        tb.append(el)


def bordas_tabela(t, cor=BORDA, sz=4, internas_verticais=False):
    tblPr = t._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{lado}")
        if lado in ("left", "right") or (lado == "insideV" and not internas_verticais):
            el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:color"), cor)
        b.append(el)
    tblPr.append(b)


def margens_celula(t, cima=60, baixo=60, esq=90, dir_=90):
    tblPr = t._tbl.tblPr
    m = OxmlElement("w:tblCellMar")
    for lado, v in (("top", cima), ("left", esq), ("bottom", baixo), ("right", dir_)):
        el = OxmlElement(f"w:{lado}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        m.append(el)
    tblPr.append(m)


def repetir_cabecalho(linha):
    trPr = linha._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def nao_quebrar_linha(linha):
    trPr = linha._tr.get_or_add_trPr()
    el = OxmlElement("w:cantSplit")
    el.set(qn("w:val"), "true")
    trPr.append(el)


def borda_paragrafo(par, lado="bottom", cor=TEAL, sz=12, espaco=4):
    pPr = par._p.get_or_add_pPr()
    pb = OxmlElement("w:pBdr")
    el = OxmlElement(f"w:{lado}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(espaco))
    el.set(qn("w:color"), cor)
    pb.append(el)
    pPr.append(pb)


def campo(par, instrucao, texto_padrao=""):
    """Insere um campo do Word (PAGE, NUMPAGES, TOC...)."""
    r = par.add_run()
    ini = OxmlElement("w:fldChar")
    ini.set(qn("w:fldCharType"), "begin")
    r._r.append(ini)
    r2 = par.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instrucao
    r2._r.append(it)
    r3 = par.add_run()
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    r3._r.append(sep)
    r4 = par.add_run(texto_padrao)
    r5 = par.add_run()
    fim = OxmlElement("w:fldChar")
    fim.set(qn("w:fldCharType"), "end")
    r5._r.append(fim)
    return r4


# ------------------------------------------------------------------ helpers de conteúdo
def marcar(par, texto, tamanho=None):
    r = par.add_run(texto)
    r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    r.bold = True
    if tamanho:
        r.font.size = Pt(tamanho)
    return r


def p(doc, texto="", negrito=False, italico=False, estilo=None, alinhamento=None, tamanho=None, cor=None,
      depois=None):
    """Parágrafo; trechos entre [[...]] viram marcação amarela da equipe."""
    par = doc.add_paragraph(style=estilo)
    partes = texto.split("[[")

    def run(t):
        r = par.add_run(t)
        r.bold = negrito
        r.italic = italico
        if tamanho:
            r.font.size = Pt(tamanho)
        if cor:
            r.font.color.rgb = cor
        return r

    run(partes[0])
    for parte in partes[1:]:
        dentro, fora = parte.split("]]", 1)
        marcar(par, f"[{dentro}]", tamanho)
        run(fora)
    if alinhamento is not None:
        par.alignment = alinhamento
    if depois is not None:
        par.paragraph_format.space_after = Pt(depois)
    return par


def formula(doc, texto):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_before = Pt(4)
    par.paragraph_format.space_after = Pt(8)
    r = par.add_run(texto)
    r.font.name = "Cambria Math"
    r.font.size = Pt(11.5)
    r.font.color.rgb = AZUL_RGB
    return par


def legenda(doc, texto, tipo="Tabela"):
    """Legenda acima de tabelas / abaixo de figuras, com numeração automática."""
    par = doc.add_paragraph(style="Caption")
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER if tipo == "Figura" else WD_ALIGN_PARAGRAPH.LEFT
    r = par.add_run(f"{tipo} ")
    campo(par, f" SEQ {tipo} \\* ARABIC ", "1")
    par.add_run(f" — {texto}")
    for r in par.runs:
        r.font.size = Pt(9)
        r.font.color.rgb = CINZA_TXT
        r.bold = True
    par.paragraph_format.keep_with_next = tipo != "Figura"
    return par


def fonte_nota(doc, texto):
    par = p(doc, texto, tamanho=8, cor=CINZA_TXT, italico=True, depois=10)
    return par


def caixa(doc, titulo, texto, tipo="pendencia"):
    """Caixa de destaque: 'pendencia' (amarela) ou 'info' (azul clara)."""
    fundo, borda = (AMARELO_BOX, AMARELO_BORDA) if tipo == "pendencia" else ("E8F4F2", TEAL)
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cel = t.rows[0].cells[0]
    sombrear(cel, fundo)
    bordas_celula(cel, left=(24, borda), top=None, bottom=None, right=None)
    margens_celula(t, 100, 100, 160, 160)
    cel.text = ""
    par = cel.paragraphs[0]
    r = par.add_run(titulo)
    r.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0x8A, 0x6D, 0x00) if tipo == "pendencia" else TEAL_RGB
    par2 = cel.add_paragraph()
    partes = texto.split("[[")
    par2.add_run(partes[0]).font.size = Pt(10)
    for parte in partes[1:]:
        dentro, fora = parte.split("]]", 1)
        marcar(par2, f"[{dentro}]", 10)
        par2.add_run(fora).font.size = Pt(10)
    for q in cel.paragraphs:
        q.paragraph_format.space_after = Pt(2)
    espaco(doc, 6)
    return t


def espaco(doc, pts=6):
    par = doc.add_paragraph()
    par.paragraph_format.space_after = Pt(0)
    par.paragraph_format.space_before = Pt(0)
    par.paragraph_format.line_spacing = Pt(pts)
    return par


def tabela(doc, cabecalho, linhas, larguras=None, tamanho=9, alinhar=None, destaque_linha=None,
           estilos_celula=None, negrito_coluna0=True):
    """Tabela com cabeçalho azul, linhas zebradas e números alinhados à direita.

    alinhar: lista 'l' / 'c' / 'r' por coluna (padrão: 1ª à esquerda, resto ao centro).
    destaque_linha: índice da linha de dados a destacar (ex.: vencedor).
    estilos_celula: dict {(linha, coluna): cor_fundo} para destacar pivôs etc.
    """
    ncol = len(cabecalho)
    alinhar = alinhar or (["l"] + ["c"] * (ncol - 1))
    mapa = {"l": WD_ALIGN_PARAGRAPH.LEFT, "c": WD_ALIGN_PARAGRAPH.CENTER, "r": WD_ALIGN_PARAGRAPH.RIGHT}
    t = doc.add_table(rows=1, cols=ncol)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    bordas_tabela(t)
    margens_celula(t)
    cab = t.rows[0]
    repetir_cabecalho(cab)
    for i, h in enumerate(cabecalho):
        cel = cab.cells[i]
        sombrear(cel, AZUL)
        cel.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        cel.text = ""
        par = cel.paragraphs[0]
        par.alignment = mapa[alinhar[i]]
        run = par.add_run(str(h))
        run.bold = True
        run.font.size = Pt(tamanho)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for k, linha in enumerate(linhas):
        row = t.add_row()
        nao_quebrar_linha(row)
        for i, v in enumerate(linha):
            cel = row.cells[i]
            cel.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            fundo = None
            if k % 2 == 1:
                fundo = ZEBRA
            if destaque_linha is not None and k == destaque_linha:
                fundo = "D7EFEC"
            if estilos_celula and (k, i) in estilos_celula:
                fundo = estilos_celula[(k, i)]
            if fundo:
                sombrear(cel, fundo)
            cel.text = ""
            par = cel.paragraphs[0]
            par.alignment = mapa[alinhar[i]]
            texto = str(v)
            if texto.startswith("[[") and texto.endswith("]]"):
                marcar(par, f"[{texto[2:-2]}]", tamanho)
                continue
            run = par.add_run(texto)
            run.font.size = Pt(tamanho)
            if (i == 0 and negrito_coluna0) or (destaque_linha is not None and k == destaque_linha):
                run.bold = True
            if estilos_celula and estilos_celula.get((k, i)) == PIVO_CEL:
                run.bold = True
    for row in t.rows:
        for cel in row.cells:
            for par in cel.paragraphs:
                par.paragraph_format.space_after = Pt(0)
                par.paragraph_format.space_before = Pt(0)
    if len(t.rows) <= 12:  # tabela curta: não deixar quebrar entre páginas
        for row in t.rows[:-1]:
            for cel in row.cells:
                for par in cel.paragraphs:
                    par.paragraph_format.keep_with_next = True
    if larguras:
        t.autofit = False
        for row in t.rows:
            for i, w in enumerate(larguras):
                row.cells[i].width = Cm(w)
    espaco(doc, 8)
    return t


def cartoes(doc, itens):
    """Linha de 'cartões' de indicadores: [(rótulo, valor, detalhe)]."""
    t = doc.add_table(rows=1, cols=len(itens))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    margens_celula(t, 120, 120, 140, 140)
    for i, (rotulo, valor, detalhe) in enumerate(itens):
        cel = t.rows[0].cells[i]
        sombrear(cel, "F2F7F6")
        bordas_celula(cel, top=(24, TEAL), left=None, right=(8, "FFFFFF"), bottom=None)
        cel.text = ""
        a = cel.paragraphs[0]
        a.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = a.add_run(rotulo.upper())
        r.font.size = Pt(7.5)
        r.bold = True
        r.font.color.rgb = CINZA_TXT
        b = cel.add_paragraph()
        b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = b.add_run(valor)
        r.font.size = Pt(17 if len(valor) <= 10 else 13)
        r.bold = True
        r.font.color.rgb = AZUL_RGB
        c = cel.add_paragraph()
        c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = c.add_run(detalhe)
        r.font.size = Pt(8)
        r.font.color.rgb = CINZA_TXT
        for q in cel.paragraphs:
            q.paragraph_format.space_after = Pt(1)
    espaco(doc, 10)


def figura(doc, caminho, texto, largura=15, altura_max=21):
    from PIL import Image
    with Image.open(caminho) as im:
        w, h = im.size
    largura = min(largura, altura_max * w / h)
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.keep_with_next = True
    par.add_run().add_picture(str(caminho), width=Cm(largura))
    legenda(doc, texto, "Figura")


def n4(x):
    return f"{x:.4f}".replace(".", ",")


def n2(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def brl(x):
    return "R$ " + n2(x)


def num(x):
    """Número da pesquisa no padrão brasileiro, sem casas inúteis."""
    if isinstance(x, (int,)) or float(x).is_integer():
        return f"{int(x):,}".replace(",", ".")
    return f"{x}".replace(".", ",")


# ------------------------------------------------------------------ cálculos
dados = pd.read_csv(BASE / "dados" / "pesquisa_mercado.csv")
dados["Custo mensal (R$)"] = (dados["USD_hora"] * 730 * CAMBIO).round(2)
dados = dados.set_index("Provedor")

M = matriz_de_triangulo(CRITERIOS, JULG)
ahp = calcular_ahp(M)
norm = normalizar_min_max(dados[CRITERIOS], TIPOS)
rk = ranking(norm, ahp["pesos"])
vencedor = rk.index[0]
segundo = rk.index[1]
sol = resolver(C, A, B)
v = sol["valores"]

# ------------------------------------------------------------------ gráficos
GRAF_DIR.mkdir(parents=True, exist_ok=True)
plt.rcParams.update({"font.family": "Arial", "font.size": 10, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.edgecolor": "#9AA5B1"})


def salvar(fig, nome):
    caminho = GRAF_DIR / nome
    fig.savefig(caminho, dpi=200, bbox_inches="tight")
    plt.close(fig)
    return caminho


# Pesos AHP
fig, ax = plt.subplots(figsize=(7, 2.6))
ordem = sorted(range(5), key=lambda i: ahp["pesos"][i])
ax.barh([CURTOS[i] for i in ordem], [ahp["pesos"][i] for i in ordem],
        color=["#" + (TEAL if i == 0 else "8FB3D9") for i in ordem], height=0.6)
for y, i in enumerate(ordem):
    ax.text(ahp["pesos"][i] + 0.005, y, f"{ahp['pesos'][i]*100:.1f}%".replace(".", ","), va="center", fontsize=9)
ax.set_xlim(0, max(ahp["pesos"]) * 1.18)
ax.xaxis.set_visible(False)
ax.spines["bottom"].set_visible(False)
g_pesos = salvar(fig, "pesos_ahp.png")

# Ranking
fig, ax = plt.subplots(figsize=(7, 3))
nomes = list(rk.index)[::-1]
vals = list(rk["Pontuação"])[::-1]
ax.barh(nomes, vals, color=["#" + (TEAL if n == vencedor else "B8C4D1") for n in nomes], height=0.6)
for y, val in enumerate(vals):
    ax.text(val + 0.01, y, n4(val), va="center", fontsize=9)
ax.set_xlim(0, 1.05)
ax.xaxis.set_visible(False)
ax.spines["bottom"].set_visible(False)
g_rank = salvar(fig, "ranking.png")

# Uso de recursos
fig, ax = plt.subplots(figsize=(7, 2.4))
usos = [float((B[i] - v[f"X{3+i}"]) / B[i] * 100) for i in range(3)]
ax.barh(RECURSOS[::-1], [100] * 3, color="#E6EBF1", height=0.55)
ax.barh(RECURSOS[::-1], usos[::-1], color=["#" + (TEAL if u >= 99.99 else "8FB3D9") for u in usos[::-1]], height=0.55)
for y, u in enumerate(usos[::-1]):
    ax.text(101, y, f"{u:.1f}% usado".replace(".", ","), va="center", fontsize=9)
ax.set_xlim(0, 125)
ax.xaxis.set_visible(False)
ax.spines["bottom"].set_visible(False)
g_rec = salvar(fig, "uso_recursos.png")

# ------------------------------------------------------------------ documento e estilos
doc = Document()
sec = doc.sections[0]
sec.page_height, sec.page_width = Cm(29.7), Cm(21)
sec.top_margin, sec.left_margin = Cm(3), Cm(3)
sec.bottom_margin, sec.right_margin = Cm(2), Cm(2)

st = doc.styles["Normal"]
st.font.name = FONTE
st.element.rPr.rFonts.set(qn("w:eastAsia"), FONTE)
st.font.size = Pt(11)
st.paragraph_format.space_after = Pt(6)
st.paragraph_format.line_spacing = 1.25

for nome, tam, antes, depois, cor in (("Heading 1", 16, 18, 8, AZUL_RGB), ("Heading 2", 13, 14, 6, AZUL_RGB),
                                       ("Heading 3", 11, 10, 4, TEAL_RGB), ("Title", 24, 0, 6, AZUL_RGB)):
    h = doc.styles[nome]
    h.font.name = FONTE
    for att in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if h.element.rPr.rFonts.get(qn(att)) is not None:
            del h.element.rPr.rFonts.attrib[qn(att)]
    h.element.rPr.rFonts.set(qn("w:ascii"), FONTE)
    h.element.rPr.rFonts.set(qn("w:hAnsi"), FONTE)
    h.font.size = Pt(tam)
    h.font.bold = True
    h.font.italic = False
    h.font.color.rgb = cor
    h.paragraph_format.space_before = Pt(antes)
    h.paragraph_format.space_after = Pt(depois)
    h.paragraph_format.keep_with_next = True

lb = doc.styles["List Bullet"]
lb.font.size = Pt(10)
lb.paragraph_format.space_after = Pt(2)


def h1(texto, nova_pagina=True):
    if nova_pagina:
        doc.add_page_break()
    par = doc.add_heading(texto, level=1)
    borda_paragrafo(par, "bottom", TEAL, 12, 4)
    return par


def h2(texto):
    return doc.add_heading(texto, level=2)


def h3(texto):
    return doc.add_heading(texto, level=3)


# ------------------------------------------------------------------ capa
faixa = doc.add_table(rows=1, cols=1)
faixa.alignment = WD_TABLE_ALIGNMENT.CENTER
cel = faixa.rows[0].cells[0]
sombrear(cel, AZUL)
margens_celula(faixa, 260, 260, 300, 300)
cel.text = ""
for i, linha in enumerate(["FATEC ARARAQUARA — CENTRO PAULA SOUZA",
                           "Curso Superior de Tecnologia em Análise e Desenvolvimento de Sistemas",
                           "Introdução à Pesquisa Operacional para Tomada de Decisão",
                           "Prof. Dr. Carlos Ricardo Bifi"]):
    par = cel.paragraphs[0] if i == 0 else cel.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_after = Pt(2)
    r = par.add_run(linha)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r.font.size = Pt(12 if i == 0 else 10)
    r.bold = i == 0

for _ in range(5):
    espaco(doc, 16)

par = doc.add_paragraph()
par.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = par.add_run("PROJETO FINAL")
r.font.size = Pt(10)
r.bold = True
r.font.color.rgb = TEAL_RGB
r.font.all_caps = True
par = doc.add_paragraph(style="Title")
par.alignment = WD_ALIGN_PARAGRAPH.CENTER
par.add_run("Sistema de Apoio à Decisão para Infraestrutura em Nuvem")
par = doc.add_paragraph()
par.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = par.add_run("Seleção de provedor por AHP e dimensionamento ótimo por Programação Linear")
r.font.size = Pt(13)
r.font.color.rgb = CINZA_TXT
borda_paragrafo(par, "bottom", TEAL, 12, 10)

for _ in range(4):
    espaco(doc, 16)

par = doc.add_paragraph()
par.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = par.add_run("INTEGRANTES")
r.bold = True
r.font.size = Pt(9)
r.font.color.rgb = CINZA_TXT
for _ in range(3):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    marcar(par, "[PREENCHER: Nome completo — RA]")

for _ in range(6):
    espaco(doc, 16)
par = doc.add_paragraph("Araraquara — novembro de 2026")
par.alignment = WD_ALIGN_PARAGRAPH.CENTER
par.runs[0].font.color.rgb = CINZA_TXT

# ------------------------------------------------------------------ sumário (seção nova, com cabeçalho/rodapé)
corpo = doc.add_section(WD_SECTION.NEW_PAGE)
corpo.header.is_linked_to_previous = False
corpo.footer.is_linked_to_previous = False
doc.sections[0].different_first_page_header_footer = True

hp = corpo.header.paragraphs[0]
hp.text = ""
r = hp.add_run("Projeto Final de Pesquisa Operacional  ·  Apoio à Decisão em Nuvem")
r.font.size = Pt(8)
r.font.color.rgb = CINZA_TXT
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
borda_paragrafo(hp, "bottom", BORDA, 4, 4)

fp = corpo.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
for rr in (fp.add_run("Página "),):
    rr.font.size = Pt(8)
campo(fp, " PAGE ", "1").font.size = Pt(8)
fp.add_run(" de ").font.size = Pt(8)
campo(fp, " NUMPAGES ", "1").font.size = Pt(8)
for rr in fp.runs:
    rr.font.color.rgb = CINZA_TXT
    rr.font.size = Pt(8)

par = doc.add_paragraph()
r = par.add_run("Sumário")
r.bold = True
r.font.size = Pt(16)
r.font.color.rgb = AZUL_RGB
borda_paragrafo(par, "bottom", TEAL, 12, 4)
toc = doc.add_paragraph()
campo(toc, ' TOC \\o "1-2" \\h \\z \\u ', "Clique com o botão direito aqui e escolha “Atualizar campo” para montar o sumário.")

# ------------------------------------------------------------------ sumário executivo
h1("Sumário Executivo")
cartoes(doc, [
    ("Provedor recomendado", vencedor, f"pontuação AHP {n4(rk.iloc[0]['Pontuação'])}"),
    ("Infraestrutura ótima", f"{fmt(v['X1'])} + {fmt(v['X2'])}", "microsserviços + nós analíticos"),
    ("Capacidade máxima", f"{n2(float(sol['z'])).split(',')[0]} req/s", "valor ótimo de Z"),
    ("Consistência AHP", f"RC = {n4(ahp['rc'])}", "abaixo do limite de 0,10"),
])
p(doc, "Este trabalho apoia a decisão de uma startup de TI que precisa escolher um provedor de computação em nuvem e "
       "dimensionar sua infraestrutura de produção. Na Etapa 1, seis provedores (AWS, Microsoft Azure, Google Cloud, "
       "Oracle Cloud, DigitalOcean e Linode/Akamai) foram comparados em uma instância equivalente de 2 vCPU e 8 GB de "
       "RAM, sob cinco critérios ponderados pelo método AHP de Saaty, após normalização linear Min-Max. "
       f"O provedor recomendado foi {vencedor}, com pontuação global de {n4(rk.iloc[0]['Pontuação'])}.")
p(doc, "Na Etapa 2, os preços do provedor vencedor alimentaram um modelo de Programação Linear resolvido pelo Simplex "
       f"Tabular, que indicou {fmt(v['X1'])} instâncias de microsserviços e {fmt(v['X2'])} nós analíticos, atingindo "
       f"{n2(float(sol['z'])).split(',')[0]} requisições por segundo. Todo o processo foi automatizado em um dashboard "
       "interativo desenvolvido em Python/Streamlit.")
caixa(doc, "PARA A EQUIPE", "[[CONFERIR: revisar este resumo antes da entrega]]")

# ------------------------------------------------------------------ 1. contextualização
h1("1. Contextualização do Cenário em ADS")
p(doc, "A [[PREENCHER: nome da startup fictícia]] é uma startup de Araraquara que desenvolve uma plataforma SaaS de "
       "gestão para pequenos comércios. A plataforma tem duas cargas de trabalho:")
p(doc, "Microsserviços transacionais, que atendem as requisições dos usuários em tempo real;", estilo="List Bullet")
p(doc, "Nós de processamento analítico, que geram relatórios e indicadores de vendas.", estilo="List Bullet")
p(doc, "O problema de negócio tem duas decisões encadeadas:")
tabela(doc, ["Etapa", "Pergunta de negócio", "Método", "Saída"], [
    ["1", "Em qual provedor hospedar a plataforma?", "AHP + Min-Max", "Provedor e seus preços"],
    ["2", "Quantas instâncias de cada tipo contratar?", "PL + Simplex", "Mix ótimo de instâncias"],
], larguras=[1.4, 6.6, 3.2, 4.8], alinhar=["c", "l", "l", "l"])
p(doc, "A saída da Etapa 1 (o provedor e seus preços) é a entrada da Etapa 2 (os coeficientes de custo do modelo de PL). "
       "Assim, a equipe equilibra custo, desempenho e confiabilidade sem estourar a cota de vCPU, a memória disponível "
       "e o orçamento mensal.")

# ------------------------------------------------------------------ 2. etapa 1
h1("2. Etapa 1 — Estudo Multicritério (AHP e Normalização)")
h2("2.1 Alternativas pesquisadas")
p(doc, "Para comparar alternativas equivalentes, fixou-se uma instância de 2 vCPU e 8 GB de RAM, Linux, sob demanda, "
       "na região de São Paulo sempre que o provedor a oferece. O preço mensal considera 730 horas e câmbio de "
       "R$ 5,4701 por dólar (cotação de compra da AwesomeAPI em 04/11/2026 às 14h18).")
formula(doc, "Custo mensal = US$/hora × 730 h × 5,4701")
legenda(doc, "Instâncias equivalentes comparadas")
tabela(doc, ["Provedor", "Instância", "Região", "US$/h", "R$/mês"],
       [[i, r["Instância"], r["Região"], f"{r['USD_hora']:.4f}".replace(".", ","), brl(r["Custo mensal (R$)"])]
        for i, r in dados.iterrows()],
       larguras=[2.8, 4.8, 3.9, 1.7, 2.8], alinhar=["l", "l", "l", "r", "r"])
legenda(doc, "Fontes consultadas e disco considerado")
tabela(doc, ["Provedor", "Fonte do preço", "Fonte do SLA", "Disco considerado"],
       [[i, r["Fonte_preco"], r["Fonte_SLA"], r["Disco"]] for i, r in dados.iterrows()],
       larguras=[2.3, 5.2, 4.5, 4], tamanho=7, alinhar=["l", "l", "l", "l"])
caixa(doc, "PARA A EQUIPE", "[[CONFERIR: anexar prints das fontes com data]]")

h2("2.2 Critérios de avaliação")
legenda(doc, "Critérios, sentido de otimização e justificativa")
tabela(doc, ["Critério", "Tipo", "Unidade", "Justificativa técnica"], [
    ["C1 Custo mensal", "Custo ↓", "R$/mês", "Determina o gasto recorrente (OPEX) e a viabilidade da startup."],
    ["C2 Latência SP", "Custo ↓", "ms", "Tempo até o datacenter; afeta a experiência do usuário final."],
    ["C3 SLA", "Benefício ↑", "%", "Disponibilidade garantida em contrato para uma VM única; reduz o risco de indisponibilidade."],
    ["C4 IOPS", "Benefício ↑", "operações/s", "Desempenho do disco padrão; limita bancos de dados e filas."],
    ["C5 Suporte", "Benefício ↑", "nota 1 a 5", "Qualidade do suporte técnico, medida por uma régua própria."],
], larguras=[3, 2.2, 2.3, 8.5], alinhar=["l", "c", "c", "l"])
p(doc, "Régua de suporte (1 ponto por item atendido):")
for item in ["atendimento 24/7;", "suporte técnico no plano gratuito;", "atendimento em português;",
             "chat ou telefone;", "tempo de resposta garantido."]:
    p(doc, item, estilo="List Bullet")
caixa(doc, "PARA A EQUIPE", "[[CONFERIR: pontuar cada provedor na régua com fonte]]")

legenda(doc, "Matriz de decisão (valores brutos)")
tabela(doc, ["Provedor", "Custo (R$/mês)", "Latência (ms)", "SLA (%)", "IOPS", "Suporte"],
       [[i, n2(r[CRITERIOS[0]]), num(r[CRITERIOS[1]]), num(r[CRITERIOS[2]]), num(r[CRITERIOS[3]]), num(r[CRITERIOS[4]])]
        for i, r in dados.iterrows()],
       larguras=[3.2, 2.8, 2.6, 2.2, 2.4, 2.4], alinhar=["l", "r", "r", "r", "r", "c"])
fonte_nota(doc, "Latência: mediana de 10 medições do tempo de conexão TCP na porta 443 (curl), feitas de Araraquara em "
                "04/11/2026 contra um endpoint de cada região (AWS: ec2.sa-east-1.amazonaws.com; Azure: "
                "s3brazilsouth.blob.core.windows.net; Google: endpoint gcping de southamerica-east1; Oracle: "
                "objectstorage.sa-saopaulo-1.oraclecloud.com; DigitalOcean: nyc3.digitaloceanspaces.com; Linode: "
                "br-gru-1.linodeobjects.com). Usou-se TCP porque AWS e Azure bloqueiam ping.")

h2("2.3 Normalização Min-Max")
p(doc, "Para que critérios com unidades diferentes possam ser somados, cada coluna é levada para a escala [0, 1]. "
       "O melhor valor recebe 1 e o pior recebe 0:")
formula(doc, "Benefício:  Nᵢⱼ = (Xᵢⱼ − Xⱼ,mín) / (Xⱼ,máx − Xⱼ,mín)")
formula(doc, "Custo:  Nᵢⱼ = (Xⱼ,máx − Xᵢⱼ) / (Xⱼ,máx − Xⱼ,mín)")
legenda(doc, "Extremos de cada critério")
tabela(doc, ["Critério", "Tipo", "Mínimo", "Máximo", "Amplitude"],
       [[c, "Custo ↓" if TIPOS[c] == "custo" else "Benefício ↑", num(dados[c].min()), num(dados[c].max()),
         num(round(dados[c].max() - dados[c].min(), 4))] for c in CRITERIOS],
       larguras=[4.2, 2.8, 2.8, 2.8, 2.8], alinhar=["l", "c", "r", "r", "r"])
ex = rk.index[min(2, len(rk) - 1)]
c0 = CRITERIOS[0]
caixa(doc, "MEMÓRIA DE CÁLCULO", f"{ex}, {c0} (critério de custo):  N = ({n2(dados[c0].max())} − "
      f"{n2(dados.loc[ex, c0])}) / ({n2(dados[c0].max())} − {n2(dados[c0].min())}) = {n4(norm.loc[ex, c0])}", "info")
legenda(doc, "Matriz normalizada Nᵢⱼ")
tabela(doc, ["Provedor"] + CURTOS, [[i] + [n4(x) for x in r] for i, r in norm.iterrows()],
       larguras=[3.2, 2.56, 2.56, 2.56, 2.56, 2.56], alinhar=["l"] + ["r"] * 5)

h2("2.4 Matriz de Saaty e pesos")
p(doc, "Julgamentos da equipe: o custo é o critério mais importante para uma startup em fase inicial (moderadamente mais "
       "importante que latência e SLA, fortemente mais importante que IOPS e suporte). Latência e SLA têm a mesma "
       "importância entre si e são moderadamente mais importantes que IOPS e suporte, que empatam.")
caixa(doc, "PARA A EQUIPE", "[[CONFERIR: a equipe deve revisar e defender cada julgamento]]")


def frac(x):
    for d in (1, 3, 5, 7, 9, 2, 4, 6, 8):
        if abs(x * d - round(x * d)) < 1e-9:
            n_ = round(x * d)
            return str(n_) if d == 1 else f"{n_}/{d}"
    return f"{x:.3f}".replace(".", ",")


legenda(doc, "Matriz de comparação par a par A (escala de Saaty)")
diag = {(k, k + 1): "EAF0F7" for k in range(5)}
tabela(doc, [""] + CURTOS, [[CURTOS[i]] + [frac(x) for x in M[i]] for i in range(5)]
       + [["Soma"] + [n4(x) for x in ahp["soma_colunas"]]],
       larguras=[3.2] + [2.56] * 5, alinhar=["l"] + ["c"] * 5, estilos_celula=diag)
legenda(doc, "Matriz normalizada (÷ soma da coluna) e vetor de pesos W")
tabela(doc, [""] + CURTOS + ["Peso W"],
       [[CURTOS[i]] + [n4(x) for x in ahp["normalizada"][i]] + [n4(ahp["pesos"][i])] for i in range(5)],
       larguras=[2.6] + [2.2] * 5 + [2.4], alinhar=["l"] + ["r"] * 6,
       estilos_celula={(i, 6): "D7EFEC" for i in range(5)})
figura(doc, g_pesos, "Pesos dos critérios obtidos pelo AHP", 13)

h3("Verificação de consistência")
legenda(doc, "Cálculo de λmax")
tabela(doc, ["Critério", "(A·w)ᵢ", "wᵢ", "(A·w)ᵢ / wᵢ"],
       [[CURTOS[i], n4(ahp["aw"][i]), n4(ahp["pesos"][i]), n4(ahp["aw"][i] / ahp["pesos"][i])] for i in range(5)],
       larguras=[4, 3.5, 3.5, 3.5], alinhar=["l", "r", "r", "r"])
formula(doc, f"λmax = {n4(ahp['lambda_max'])}     IC = (λmax − 5) / 4 = {n4(ahp['ic'])}     "
             f"RC = IC / 1,12 = {n4(ahp['rc'])}")
caixa(doc, "RESULTADO", ("Como RC < 0,10, os julgamentos são consistentes." if ahp["consistente"]
                         else "RC ≥ 0,10: revisar julgamentos."), "info")

h2("2.5 Ranking final e parecer")
formula(doc, "Pontuaçãoᵢ = Σⱼ Wⱼ · Nᵢⱼ")
legenda(doc, "Ranking final dos provedores")
tabela(doc, ["Posição", "Provedor", "Pontuação"],
       [[f"{int(r['Posição'])}º", i, n4(r["Pontuação"])] for i, r in rk.iterrows()],
       larguras=[2.2, 5, 3.5], alinhar=["c", "l", "r"], destaque_linha=0, negrito_coluna0=False)
figura(doc, g_rank, "Pontuação global de cada provedor", 13)
p(doc, f"Parecer: {vencedor} obteve a maior pontuação, à frente de {segundo}.", negrito=True)
caixa(doc, "PARA A EQUIPE", "[[PREENCHER: explicar em quais critérios o vencedor se destacou, por que superou o 2º "
                            "colocado e quais riscos a equipe aceita ao escolhê-lo]]")

# ------------------------------------------------------------------ 3. etapa 2
h1("3. Etapa 2 — Programação Linear e Método Simplex")
h2("3.1 Variáveis, função objetivo e restrições")
p(doc, f"Com o provedor {vencedor} escolhido, a equipe dimensiona a infraestrutura de produção:")
legenda(doc, "Variáveis de decisão")
tabela(doc, ["Variável", "Significado", "Configuração", "Capacidade", "Custo/mês"], [
    ["X1", "Microsserviços Transacionais", "2 vCPU · 8 GB", "500 req/s", "≈ R$ 148"],
    ["X2", "Nós de Processamento Analítico", "4 vCPU · 32 GB", "1.200 req/s", "≈ R$ 391"],
], larguras=[1.8, 5.2, 3, 2.8, 2.8], alinhar=["c", "l", "c", "r", "r"])
caixa(doc, "PARA A EQUIPE", "[[CONFERIR: justificar a capacidade em req/s de cada tipo, por exemplo com um teste de "
                            "carga ou benchmark]]")
formula(doc, f"Max Z = {C[0]}X₁ + {num(C[1])}X₂")
p(doc, "Sujeito a:", alinhamento=WD_ALIGN_PARAGRAPH.CENTER, cor=CINZA_TXT, tamanho=10, depois=2)
for r_, a, b_ in zip(RECURSOS, A, B):
    formula(doc, f"{a[0]}X₁ + {a[1]}X₂ ≤ {num(b_)}     ({r_})").paragraph_format.space_after = Pt(2)
formula(doc, "X₁, X₂ ≥ 0")
p(doc, f"Os custos vêm do preço do vencedor (API oficial da Oracle: US$ 0,025 por OCPU-hora e US$ 0,0015 por GB-hora): "
       "X1 = 1 OCPU + 8 GB ≈ US$ 0,037/h ≈ R$ 148/mês; X2 = 2 OCPU + 32 GB = US$ 0,098/h ≈ R$ 391/mês. "
       "A cota de 64 vCPU é o limite da conta; 320 GB é o limite de memória contratado; R$ 6.000 é a verba mensal de TI.")

h2("3.2 Forma padrão")
p(doc, "Somando uma variável de folga a cada restrição, as desigualdades viram igualdades:")
formula(doc, f"Max Z = {C[0]}X₁ + {num(C[1])}X₂ + 0X₃ + 0X₄ + 0X₅")
for i, (a, b_) in enumerate(zip(A, B)):
    formula(doc, f"{a[0]}X₁ + {a[1]}X₂ + X{'₃₄₅'[i]} = {num(b_)}").paragraph_format.space_after = Pt(2)
formula(doc, "X₁, X₂, X₃, X₄, X₅ ≥ 0")
legenda(doc, "Significado das variáveis de folga")
tabela(doc, ["Folga", "Significado físico no sistema de TI"],
       [[f"X{3+i}", s] for i, s in enumerate(SIGNIF_FOLGA)], larguras=[2.2, 10], alinhar=["c", "l"])
p(doc, "As colunas de X3, X4 e X5 formam a Matriz Identidade inicial, por isso as folgas são a base do primeiro quadro.")

h2("3.3 Iterações do Simplex Tabular")
cx = doc.add_table(rows=1, cols=3)
cx.alignment = WD_TABLE_ALIGNMENT.CENTER
margens_celula(cx, 40, 40, 100, 100)
for k, (cor, txt) in enumerate([(PIVO_COL, "coluna que entra"), (PIVO_LIN, "linha que sai"), (PIVO_CEL, "elemento pivô")]):
    cel = cx.rows[0].cells[k]
    sombrear(cel, cor)
    cel.text = ""
    rr = cel.paragraphs[0].add_run(f"Legenda: {txt}" if k == 0 else txt)
    rr.font.size = Pt(8)
    cel.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    cel.paragraphs[0].paragraph_format.space_after = Pt(0)
espaco(doc, 8)

for passo in sol["passos"]:
    nome = ("Quadro inicial (Acampamento Base, Z = 0)" if passo["iteracao"] == 0
            else f"Quadro após a iteração {passo['iteracao']}")
    if passo["status"] == "otimo":
        nome += " — solução ótima"
    h3(nome)
    variaveis = passo["variaveis"]
    cab = ["Base", "Cb"] + variaveis + ["RHS"]
    tem_razao = passo["status"] == "continua"
    if tem_razao:
        cab.append("Razão")
    linhas = []
    for i, bv in enumerate(passo["base"]):
        lin = [bv, fmt(passo["cb"][i])] + [fmt(x) for x in passo["tabela"][i]]
        if tem_razao:
            rz = passo["razoes"][i]
            lin.append(fmt(rz) if rz is not None else "—")
        linhas.append(lin)
    nb = len(linhas)
    linhas.append(["Zj", ""] + [fmt(x) for x in passo["zj"]] + [fmt(passo["z"])] + ([""] if tem_razao else []))
    linhas.append(["Cj − Zj", ""] + [fmt(x) for x in passo["cj_zj"]] + [""] + ([""] if tem_razao else []))
    estilos = {(nb, k): "E6EBF1" for k in range(len(cab))}
    estilos.update({(nb + 1, k): "E6EBF1" for k in range(len(cab))})
    if tem_razao:
        col = passo["col_pivo"] + 2
        lin = passo["lin_pivo"]
        for k in range(nb + 2):
            estilos[(k, col)] = PIVO_COL
        for k in range(len(cab)):
            estilos[(lin, k)] = PIVO_LIN
        estilos[(lin, col)] = PIVO_CEL
    if passo["status"] == "otimo":
        estilos[(nb, len(cab) - 1)] = "D7EFEC"
    ncol = len(cab)
    larg_total = 16
    larg = [1.5, 1.3] + [(larg_total - 2.8) / (ncol - 2)] * (ncol - 2)
    tabela(doc, cab, linhas, larguras=larg, tamanho=8.5, alinhar=["l"] + ["c"] * (ncol - 1), estilos_celula=estilos)
    if tem_razao:
        p(doc, f"Entra {passo['entra']} (maior Cj − Zj = {fmt(passo['cj_zj'][passo['col_pivo']])}). "
               f"Sai {passo['sai']} (menor razão = {fmt(passo['razoes'][passo['lin_pivo']])}, ou seja, é o recurso que "
               f"se esgota primeiro). Elemento pivô = {fmt(passo['pivo'])}.", tamanho=10)
        p(doc, "Operações de Gauss-Jordan:", negrito=True, tamanho=10, depois=2)
        for op in passo["operacoes"]:
            p(doc, op, estilo="List Bullet")
    else:
        caixa(doc, "CRITÉRIO DE PARADA", "Todos os Cj − Zj ≤ 0, então a solução é ótima.", "info")

h2("3.4 Leitura gerencial")
cartoes(doc, [
    ("X1 · Microsserviços", fmt(v["X1"]), "instâncias"),
    ("X2 · Nós analíticos", fmt(v["X2"]), "instâncias"),
    ("Z máximo", n2(float(sol["z"])).split(",")[0], "requisições por segundo"),
])
p(doc, f"Recomendação: contratar {fmt(v['X1'])} instâncias de microsserviços e {fmt(v['X2'])} nós analíticos, "
       f"atingindo Z = {n2(float(sol['z'])).split(',')[0]} requisições por segundo.", negrito=True)
linhas = []
for i, rec in enumerate(RECURSOS):
    folga = v[f"X{3+i}"]
    linhas.append([rec, num(B[i]), num(float(B[i] - folga)), num(float(folga)),
                   "Gargalo" if folga == 0 else "Capacidade ociosa"])
legenda(doc, "Uso dos recursos na solução ótima")
tabela(doc, ["Recurso", "Disponível", "Usado", "Folga", "Situação"], linhas,
       larguras=[4.2, 2.6, 2.6, 2.6, 3.5], alinhar=["l", "r", "r", "r", "c"],
       estilos_celula={(i, 4): ("FBE3D6" if l[4] == "Gargalo" else "D7EFEC") for i, l in enumerate(linhas)})
figura(doc, g_rec, "Percentual de cada recurso consumido pela solução ótima", 13)
p(doc, "vCPU e RAM são os gargalos: para crescer, a startup precisa ampliar a cota de vCPU e a memória, não o orçamento. "
       f"Sobram {brl(float(v['X5']))} por mês, que podem ir para backup, monitoramento ou reserva.")
caixa(doc, "PARA A EQUIPE", "[[CONFERIR: ajustar este texto se os números mudarem]]")

# ------------------------------------------------------------------ 4. dashboard
h1("4. Dashboard Utilitário")
p(doc, "A ferramenta foi desenvolvida em Python com Streamlit, Pandas, NumPy e Plotly. O motor do Simplex foi "
       "implementado pela equipe com frações exatas (módulo fractions), o que evita erros de arredondamento no "
       "pivoteamento. O resultado é validado automaticamente com scipy.optimize.linprog.")
legenda(doc, "Estrutura do código")
tabela(doc, ["Arquivo", "Responsabilidade"], [
    ["app.py", "Interface com as abas Dados, AHP e Simplex"],
    ["motores/ahp.py", "Normalização Min-Max, pesos de Saaty, λmax, IC, RC e ranking"],
    ["motores/simplex.py", "Simplex tabular com registro de cada quadro e das operações de Gauss-Jordan"],
    ["motores/precos.py", "Consulta de preço ao vivo na Azure Retail Prices API"],
    ["dados/pesquisa_mercado.csv", "Pesquisa de mercado com fontes"],
    ["gerar_relatorio.py", "Gera este relatório a partir dos mesmos motores"],
], larguras=[5, 11], alinhar=["l", "l"])
caixa(doc, "COMO EXECUTAR", "pip install -r requirements.txt   →   streamlit run app.py", "info")
p(doc, "Repositório: [[PREENCHER: link do GitHub]]")
for i, leg in enumerate(["Aba 1: dados de mercado", "Aba 2: AHP e ranking", "Aba 3: Simplex e painel gerencial"], 1):
    img = BASE / "prints" / f"aba{i}.png"
    if img.exists():
        if i > 1:
            doc.add_page_break()
        figura(doc, img, leg, 14)

# ------------------------------------------------------------------ 5. apêndice IA
h1("5. Apêndice — Governança e Rastreabilidade do Uso de IA")
p(doc, "A equipe usou IA generativa como copiloto. Todas as saídas foram auditadas conforme a última coluna.")
legenda(doc, "Registro de uso de IA")
tabela(doc, ["ID / Etapa", "IA e versão", "Objetivo", "Prompt / entrada", "Saída gerada", "Auditoria humana"], [
    ["#01 Planejamento", "Claude Opus 5.5 (Anthropic)", "Sugerir abordagem para o projeto",
     "\"De acordo com o arquivo da proposta do projeto, me dê ideias de como realizá-lo\"",
     "Sugestão do Cenário B, critérios e modelo de PL", "[[PREENCHER: o que a equipe aceitou ou mudou]]"],
    ["#02 Etapas 1 e 2", "Claude Opus 5.5 (Anthropic)", "Gerar o protótipo do dashboard",
     "\"Pode montar pra eu ver se é isso mesmo que quero\"",
     "app.py, motores/ahp.py, motores/simplex.py",
     "[[PREENCHER: conferência manual de uma normalização, dos pesos e de um pivoteamento]]"],
    ["#03 Etapa 1", "Claude Opus 5.5 (Anthropic)", "Levantar preços, SLA, IOPS e medir latência no PC da equipe",
     "\"Faça o que for possível no momento\"", "dados/pesquisa_mercado.csv com fontes",
     "[[PREENCHER: conferência de cada preço na calculadora oficial]]"],
    ["#04 Documentação", "Claude Opus 5.5 (Anthropic)", "Gerar o relatório e melhorar o visual",
     "Mesmo pedido do #03 e \"refaça o relatório de um jeito mais bonito\"", "gerar_relatorio.py e este documento",
     "[[PREENCHER: revisão do texto pela equipe]]"],
], larguras=[2.7, 2.2, 2.6, 3.1, 2.7, 2.7], tamanho=7.5, alinhar=["l"] * 6)
p(doc, "Validações já feitas: o Simplex próprio foi comparado com scipy.optimize.linprog (X1 = 24, X2 = 4, Z = 16.800 "
       "nos dois).")
caixa(doc, "PARA A EQUIPE", "[[PREENCHER: refazer à mão um pivoteamento e uma linha da normalização e registrar aqui]]")

SAIDA_DIR.mkdir(exist_ok=True)
saida = SAIDA_DIR / "Relatorio_Projeto_Final_PO.docx"
doc.save(saida)
print("Relatório gerado em", saida)

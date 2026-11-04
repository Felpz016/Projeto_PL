"""Gera o relatório Word (.docx) do projeto a partir dos mesmos motores do dashboard.

Rodar:  python gerar_relatorio.py
Saída:  relatorio/Relatorio_Projeto_Final_PO.docx

Trechos entre [[...]] saem destacados em amarelo (use para pendências).
"""
import io
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from motores.ahp import calcular_ahp, matriz_de_triangulo, normalizar_min_max, ranking
from motores.simplex import fmt, resolver

BASE = Path(__file__).parent

# ------------------------------------------------------------------ parâmetros do estudo
CAMBIO = 5.4701  # AwesomeAPI, 04/11/2026 14h18
CRITERIOS = ["Custo mensal (R$)", "Latência SP (ms)", "SLA (%)", "IOPS", "Suporte (1-5)"]
CURTOS = ["Custo", "Latência", "SLA", "IOPS", "Suporte"]
TIPOS = {"Custo mensal (R$)": "custo", "Latência SP (ms)": "custo",
         "SLA (%)": "beneficio", "IOPS": "beneficio", "Suporte (1-5)": "beneficio"}
JULG = {(0, 1): 3, (0, 2): 3, (0, 3): 5, (0, 4): 5, (1, 2): 1, (1, 3): 3, (1, 4): 3,
        (2, 3): 3, (2, 4): 3, (3, 4): 1}
C = [500, 1200]
A = [[2, 4], [8, 32], [148, 391]]
B = [64, 320, 6000]
RECURSOS = ["vCPU (cota)", "RAM (GB)", "Orçamento (R$/mês)"]
SIGNIF_FOLGA = ["vCPUs da cota da conta que ficam sem uso",
                "GB de memória contratada que ficam sem uso",
                "Reais do orçamento mensal de TI que não são gastos"]

# ------------------------------------------------------------------ paleta
AZUL = RGBColor(0x1F, 0x3A, 0x5F)
AZUL_HEX = "1F3A5F"
AZUL_CLARO = "E8EEF6"
CINZA = "F5F7FA"
VERDE = "E3F2E8"
AMARELO = "FFF4CC"
LARANJA = "F8C8A0"
AZUL_LINHA = "D6E6F7"
FONTE = "Calibri"


# ------------------------------------------------------------------ utilitários de formatação
def sombrear(cel, cor):
    tcPr = cel._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), cor)
    tcPr.append(shd)


def bordas_tabela(tabela, cor="BFC8D6", tam=4):
    tblPr = tabela._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(tam))
        e.set(qn("w:color"), cor)
        b.append(e)
    tblPr.append(b)


def sem_bordas(tabela):
    tblPr = tabela._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:val"), "nil")
        b.append(e)
    tblPr.append(b)


def margens_celula(tabela, cm=0.12):
    tblPr = tabela._tbl.tblPr
    m = OxmlElement("w:tblCellMar")
    for lado in ("top", "bottom", "left", "right"):
        e = OxmlElement(f"w:{lado}")
        e.set(qn("w:w"), str(int(cm * 567)))
        e.set(qn("w:type"), "dxa")
        m.append(e)
    tblPr.append(m)


def repetir_cabecalho(linha):
    trPr = linha._tr.get_or_add_trPr()
    e = OxmlElement("w:tblHeader")
    e.set(qn("w:val"), "true")
    trPr.append(e)


def borda_inferior(par, cor=AZUL_HEX, tam=12):
    pPr = par._p.get_or_add_pPr()
    pb = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), str(tam))
    bot.set(qn("w:space"), "4")
    bot.set(qn("w:color"), cor)
    pb.append(bot)
    pPr.append(pb)


def campo(par, instr, texto_inicial="1"):
    r = par.add_run()
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = instr
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = texto_inicial
    f3 = OxmlElement("w:fldChar")
    f3.set(qn("w:fldCharType"), "end")
    for e in (f1, it, f2, t, f3):
        r._r.append(e)
    return r


def texto_com_pendencias(par, texto, negrito=False, italico=False, tamanho=None, cor=None):
    """Escreve texto; trechos entre [[...]] saem destacados em amarelo."""
    partes = texto.split("[[")
    def run(t, destaque=False):
        r = par.add_run(t)
        r.bold = negrito
        r.italic = italico
        if tamanho:
            r.font.size = Pt(tamanho)
        if cor:
            r.font.color.rgb = cor
        if destaque:
            r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        return r
    run(partes[0])
    for parte in partes[1:]:
        dentro, fora = parte.split("]]", 1)
        run(f"[{dentro}]", destaque=True)
        run(fora)


class Relatorio:
    def __init__(self):
        self.doc = Document()
        self._estilos()
        self.n_tab = 0
        self.n_fig = 0

    # ---------------------------------------------------------- estilos
    def _estilos(self):
        d = self.doc
        sec = d.sections[0]
        sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
        sec.top_margin, sec.bottom_margin = Cm(3.0), Cm(2.0)
        sec.left_margin, sec.right_margin = Cm(3.0), Cm(2.0)

        st = d.styles
        normal = st["Normal"]
        normal.font.name = FONTE
        normal.element.rPr.rFonts.set(qn("w:eastAsia"), FONTE)
        normal.font.size = Pt(11)
        normal.paragraph_format.space_after = Pt(6)
        normal.paragraph_format.line_spacing = 1.25

        for nome, tam, antes, depois in (("Heading 1", 18, 18, 8), ("Heading 2", 13.5, 14, 6), ("Heading 3", 11.5, 10, 4)):
            h = st[nome]
            h.font.name = FONTE
            h.element.rPr.rFonts.set(qn("w:asciiTheme"), "") if False else None
            rfonts = h.element.rPr.find(qn("w:rFonts"))
            if rfonts is None:
                rfonts = OxmlElement("w:rFonts")
                h.element.rPr.append(rfonts)
            for k in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
                rfonts.set(qn(k), FONTE)
            for k in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
                if rfonts.get(qn(k)) is not None:
                    del rfonts.attrib[qn(k)]
            h.font.size = Pt(tam)
            h.font.bold = True
            h.font.color.rgb = AZUL
            h.paragraph_format.space_before = Pt(antes)
            h.paragraph_format.space_after = Pt(depois)
            h.paragraph_format.keep_with_next = True

        # Word atualiza o sumário ao abrir
        settings = d.settings.element
        uf = OxmlElement("w:updateFields")
        uf.set(qn("w:val"), "true")
        settings.append(uf)

    # ---------------------------------------------------------- blocos de texto
    def h1(self, texto, nova_pagina=True):
        if nova_pagina:
            self.quebra()
        par = self.doc.add_heading(texto, level=1)
        borda_inferior(par)
        return par

    def h2(self, texto):
        return self.doc.add_heading(texto, level=2)

    def h3(self, texto):
        return self.doc.add_heading(texto, level=3)

    def p(self, texto, negrito=False, italico=False, alinhar="just", tamanho=None, depois=None):
        par = self.doc.add_paragraph()
        texto_com_pendencias(par, texto, negrito, italico, tamanho)
        par.alignment = {"just": WD_ALIGN_PARAGRAPH.JUSTIFY, "centro": WD_ALIGN_PARAGRAPH.CENTER,
                         "esq": WD_ALIGN_PARAGRAPH.LEFT}[alinhar]
        if depois is not None:
            par.paragraph_format.space_after = Pt(depois)
        return par

    def lista(self, itens, numerada=False):
        for it in itens:
            par = self.doc.add_paragraph(style="List Number" if numerada else "List Bullet")
            texto_com_pendencias(par, it)
            par.paragraph_format.space_after = Pt(2)

    def quebra(self):
        self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    def caixa(self, titulo, linhas, cor=AZUL_CLARO, borda=AZUL_HEX, mono=False):
        """Caixa de destaque (fórmula, conclusão)."""
        t = self.doc.add_table(rows=1, cols=1)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        cel = t.rows[0].cells[0]
        sombrear(cel, cor)
        tblPr = t._tbl.tblPr
        b = OxmlElement("w:tblBorders")
        e = OxmlElement("w:left")
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "24")
        e.set(qn("w:color"), borda)
        b.append(e)
        for lado in ("top", "bottom", "right"):
            x = OxmlElement(f"w:{lado}")
            x.set(qn("w:val"), "nil")
            b.append(x)
        tblPr.append(b)
        margens_celula(t, 0.25)
        par = cel.paragraphs[0]
        if titulo:
            r = par.add_run(titulo)
            r.bold = True
            r.font.color.rgb = AZUL
            par.paragraph_format.space_after = Pt(3)
        for i, linha in enumerate(linhas):
            pp = par if (i == 0 and not titulo) else cel.add_paragraph()
            pp.paragraph_format.space_after = Pt(2)
            texto_com_pendencias(pp, linha)
            if mono:
                for r in pp.runs:
                    r.font.name = "Consolas"
                    r.font.size = Pt(9)
        self.doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # ---------------------------------------------------------- tabelas
    def tabela(self, legenda, cabecalho, linhas, larguras=None, tamanho=9, alinhar_num=True,
               destaques=None, primeira_negrito=True, fonte=None):
        """destaques: dict {(linha, coluna) ou (linha, None): cor}; linha 0 = 1ª linha de dados."""
        self.n_tab += 1
        cap = self.doc.add_paragraph()
        r = cap.add_run(f"Tabela {self.n_tab} – ")
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = AZUL
        r2 = cap.add_run(legenda)
        r2.font.size = Pt(9.5)
        cap.paragraph_format.space_after = Pt(3)
        cap.paragraph_format.keep_with_next = True

        t = self.doc.add_table(rows=1, cols=len(cabecalho))
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        bordas_tabela(t)
        margens_celula(t, 0.1)
        repetir_cabecalho(t.rows[0])
        for i, h in enumerate(cabecalho):
            cel = t.rows[0].cells[i]
            sombrear(cel, AZUL_HEX)
            cel.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            pp = cel.paragraphs[0]
            pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pp.paragraph_format.space_after = Pt(0)
            pp.paragraph_format.line_spacing = 1.0
            rr = pp.add_run(str(h))
            rr.bold = True
            rr.font.size = Pt(tamanho)
            rr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        destaques = destaques or {}
        for li, linha in enumerate(linhas):
            cells = t.add_row().cells
            for ci, v in enumerate(linha):
                cel = cells[ci]
                cel.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                cor = destaques.get((li, ci)) or destaques.get((li, None)) or (CINZA if li % 2 else None)
                if cor:
                    sombrear(cel, cor)
                pp = cel.paragraphs[0]
                pp.paragraph_format.space_after = Pt(0)
                pp.paragraph_format.line_spacing = 1.0
                s = str(v)
                numerico = s.replace(",", "").replace(".", "").replace("/", "").replace("-", "").replace("−", "").replace(" ", "").replace("R$", "").isdigit()
                if alinhar_num and numerico:
                    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                texto_com_pendencias(pp, s, negrito=(primeira_negrito and ci == 0) or (li, ci) in destaques and destaques[(li, ci)] == LARANJA,
                                     tamanho=tamanho)
        if larguras:
            t.autofit = False
            for i, w in enumerate(larguras):
                t.columns[i].width = Cm(w)
            for row in t.rows:
                for i, w in enumerate(larguras):
                    row.cells[i].width = Cm(w)
        if fonte:
            fp = self.doc.add_paragraph()
            rf = fp.add_run(f"Fonte: {fonte}")
            rf.italic = True
            rf.font.size = Pt(8.5)
            rf.font.color.rgb = RGBColor(0x60, 0x6B, 0x7A)
        else:
            self.doc.add_paragraph().paragraph_format.space_after = Pt(0)
        return t

    def figura(self, legenda, imagem, largura=15):
        self.n_fig += 1
        par = self.doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.keep_with_next = True
        if isinstance(imagem, (str, Path)):
            par.add_run().add_picture(str(imagem), width=Cm(largura))
        else:
            par.add_run().add_picture(imagem, width=Cm(largura))
        cap = self.doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = cap.add_run(f"Figura {self.n_fig} – ")
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = AZUL
        r2 = cap.add_run(legenda)
        r2.font.size = Pt(9.5)


# ------------------------------------------------------------------ formatação de números
def n4(x):
    return f"{x:.4f}".replace(".", ",")


def n2(x):
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def brl(x):
    return "R$ " + n2(x)


def num(x):
    if isinstance(x, float) and x.is_integer():
        x = int(x)
    return f"{x:,}".replace(",", "X").replace(".", ",").replace("X", ".") if isinstance(x, int) else str(x).replace(".", ",")


def grafico(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def estilo_grafico(ax):
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color("#BFC8D6")
    ax.spines["bottom"].set_color("#BFC8D6")
    ax.tick_params(colors="#3A4656", labelsize=9)


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
pesos = dict(zip(CRITERIOS, ahp["pesos"]))
contrib = norm.mul(ahp["pesos"], axis=1)

# ------------------------------------------------------------------ gráficos
fig, ax = plt.subplots(figsize=(7, 2.6))
ordem = sorted(range(5), key=lambda i: ahp["pesos"][i])
ax.barh([CURTOS[i] for i in ordem], [ahp["pesos"][i] * 100 for i in ordem], color="#1F3A5F", height=0.55)
for i, k in enumerate(ordem):
    ax.text(ahp["pesos"][k] * 100 + 0.6, i, f"{ahp['pesos'][k]*100:.1f}%".replace(".", ","), va="center", fontsize=9, color="#1F3A5F")
ax.set_xlabel("Peso (%)", fontsize=9, color="#3A4656")
estilo_grafico(ax)
g_pesos = grafico(fig)

fig, ax = plt.subplots(figsize=(7, 3.2))
prov = list(rk.index)[::-1]
esq = [0] * len(prov)
cores = ["#1F3A5F", "#3E6A9E", "#7FA7D1", "#C2B280", "#E39B5B"]
for j, c in enumerate(CRITERIOS):
    vals = [contrib.loc[p_, c] for p_ in prov]
    ax.barh(prov, vals, left=esq, color=cores[j], label=CURTOS[j], height=0.6)
    esq = [a + b for a, b in zip(esq, vals)]
for i, p_ in enumerate(prov):
    ax.text(esq[i] + 0.01, i, n4(rk.loc[p_, "Pontuação"]), va="center", fontsize=9, color="#1F3A5F", fontweight="bold")
ax.set_xlim(0, 1.05)
ax.set_xlabel("Pontuação global = Σ Wj · Nij", fontsize=9, color="#3A4656")
ax.legend(ncol=5, fontsize=8, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.18))
estilo_grafico(ax)
g_ranking = grafico(fig)

fig, ax = plt.subplots(figsize=(7, 2.4))
usados = [float(B[i] - v[f"X{3+i}"]) / B[i] * 100 for i in range(3)]
ax.barh(RECURSOS[::-1], usados[::-1], color="#1F3A5F", height=0.5, label="Usado")
ax.barh(RECURSOS[::-1], [100 - u for u in usados[::-1]], left=usados[::-1], color="#BFD3EA", height=0.5, label="Folga")
for i, u in enumerate(usados[::-1]):
    ax.text(u / 2, i, f"{u:.1f}%".replace(".", ","), va="center", ha="center", color="white", fontsize=9, fontweight="bold")
ax.set_xlim(0, 100)
ax.set_xlabel("% do recurso disponível", fontsize=9, color="#3A4656")
ax.legend(ncol=2, fontsize=8, frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.3))
estilo_grafico(ax)
g_recursos = grafico(fig)

# ------------------------------------------------------------------ documento
R = Relatorio()
doc = R.doc

# ---- capa
faixa = doc.add_table(rows=1, cols=1)
sem_bordas(faixa)
cel = faixa.rows[0].cells[0]
sombrear(cel, AZUL_HEX)
margens_celula(faixa, 0.5)
for i, (txt, tam, neg) in enumerate([("FATEC ARARAQUARA  ·  CENTRO PAULA SOUZA", 10, True),
                                     ("Tecnologia em Análise e Desenvolvimento de Sistemas", 10, False),
                                     ("Introdução à Pesquisa Operacional para Tomada de Decisão", 10, False)]):
    pp = cel.paragraphs[0] if i == 0 else cel.add_paragraph()
    pp.paragraph_format.space_after = Pt(1)
    rr = pp.add_run(txt)
    rr.font.size = Pt(tam)
    rr.bold = neg
    rr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

for _ in range(5):
    doc.add_paragraph()
pp = doc.add_paragraph()
rr = pp.add_run("PROJETO FINAL")
rr.font.size = Pt(11)
rr.bold = True
rr.font.color.rgb = RGBColor(0xE3, 0x9B, 0x5B)
pp = doc.add_paragraph()
pp.paragraph_format.line_spacing = 1.05
rr = pp.add_run("Sistema de Apoio à Decisão para Infraestrutura em Nuvem")
rr.font.size = Pt(28)
rr.bold = True
rr.font.color.rgb = AZUL
pp = doc.add_paragraph()
borda_inferior(pp, "E39B5B", 18)
rr = pp.add_run("Seleção de provedor pelo método AHP e dimensionamento ótimo por Programação Linear (Simplex Tabular)")
rr.font.size = Pt(13)
rr.font.color.rgb = RGBColor(0x3A, 0x46, 0x56)
for _ in range(4):
    doc.add_paragraph()

info = doc.add_table(rows=4, cols=2)
sem_bordas(info)
margens_celula(info, 0.08)
for i, (k, val) in enumerate([("Autor", "Felipe Benaglia Nardini — RA 2880482421038"),
                              ("Docente", "Prof. Dr. Carlos Ricardo Bifi"),
                              ("Repositório", "https://github.com/Felpz016/Projeto_PL"),
                              ("Local e data", "Araraquara-SP, novembro de 2026")]):
    c0, c1 = info.rows[i].cells
    c0.width, c1.width = Cm(4), Cm(12)
    r0 = c0.paragraphs[0].add_run(k.upper())
    r0.bold = True
    r0.font.size = Pt(9)
    r0.font.color.rgb = RGBColor(0x60, 0x6B, 0x7A)
    texto_com_pendencias(c1.paragraphs[0], val, tamanho=11)

# ---- sumário
R.quebra()
par = doc.add_heading("Sumário", level=1)
borda_inferior(par)
toc = doc.add_paragraph()
campo(toc, 'TOC \\o "1-2" \\h \\z \\u',
      "Abra no Word e aceite atualizar os campos (ou clique aqui com o botão direito → Atualizar campo) para gerar o sumário.")

# ---- cabeçalho e rodapé (a partir da 2ª seção)
nova = doc.add_section(WD_SECTION.NEW_PAGE)
nova.header.is_linked_to_previous = False
nova.footer.is_linked_to_previous = False
hp = nova.header.paragraphs[0]
hr = hp.add_run("Projeto Final · Pesquisa Operacional · Fatec Araraquara")
hr.font.size = Pt(8.5)
hr.font.color.rgb = RGBColor(0x60, 0x6B, 0x7A)
borda_inferior(hp, "BFC8D6", 4)
fp = nova.footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
fr = fp.add_run("Página ")
fr.font.size = Pt(9)
rn = campo(fp, "PAGE")
rn.font.size = Pt(9)
fr2 = fp.add_run(" de ")
fr2.font.size = Pt(9)
rt = campo(fp, "NUMPAGES", "1")
rt.font.size = Pt(9)
doc.sections[0].different_first_page_header_footer = True

# ---- sumário executivo
par = doc.add_heading("Sumário Executivo", level=1)
borda_inferior(par)
R.p("A Cloud Search, uma startup de software, precisa decidir em qual provedor de nuvem hospedar sua plataforma e quantos "
    "servidores contratar. O trabalho resolve as duas decisões em sequência, com dados reais de mercado coletados em 04/11/2026.")
kp = doc.add_table(rows=2, cols=3)
sem_bordas(kp)
margens_celula(kp, 0.2)
kpis = [("Provedor recomendado", vencedor, f"pontuação AHP {n4(rk.iloc[0]['Pontuação'])}"),
        ("Configuração ótima", f"{fmt(v['X1'])} + {fmt(v['X2'])}", "microsserviços + nós analíticos"),
        ("Capacidade máxima", f"{num(int(sol['z']))} req/s", f"sobra de {brl(float(v['X5']))}/mês")]
for i, (rot, val, sub) in enumerate(kpis):
    c = kp.rows[0].cells[i]
    sombrear(c, AZUL_CLARO)
    sombrear(kp.rows[1].cells[i], AZUL_CLARO)
    r0 = c.paragraphs[0].add_run(rot.upper())
    r0.font.size = Pt(8)
    r0.bold = True
    r0.font.color.rgb = RGBColor(0x60, 0x6B, 0x7A)
    pv = c.add_paragraph()
    pv.paragraph_format.space_after = Pt(0)
    rv = pv.add_run(val)
    rv.font.size = Pt(17)
    rv.bold = True
    rv.font.color.rgb = AZUL
    rs = kp.rows[1].cells[i].paragraphs[0].add_run(sub)
    rs.font.size = Pt(9)
doc.add_paragraph()
R.p(f"Na Etapa 1, seis provedores (AWS, Microsoft Azure, Google Cloud, Oracle Cloud, DigitalOcean e Linode/Akamai) "
    f"foram comparados em uma instância equivalente de 2 vCPU e 8 GB de RAM, com cinco critérios normalizados pelo "
    f"método Min-Max e ponderados pelo AHP de Saaty (RC = {n4(ahp['rc'])}, consistente). {vencedor} ficou em 1º lugar, "
    f"principalmente pelo custo mensal mais baixo, seguido de {segundo}.")
R.p(f"Na Etapa 2, os preços de {vencedor} alimentaram um modelo de Programação Linear resolvido pelo Simplex Tabular. "
    f"A solução ótima contrata {fmt(v['X1'])} instâncias de microsserviços e {fmt(v['X2'])} nós analíticos, atingindo "
    f"{num(int(sol['z']))} requisições por segundo. vCPU e memória são os gargalos; o orçamento tem folga. "
    "Todo o processo está automatizado em um dashboard interativo em Python/Streamlit.")

# ---- 1. contexto
R.h1("1. Contextualização do Cenário")
R.h2("1.1 A empresa")
R.p("A Cloud Search é uma startup de Araraquara-SP que desenvolve uma plataforma SaaS de gestão para pequenos "
    "comércios. A plataforma tem duas cargas de trabalho distintas:")
R.lista(["Microsserviços transacionais, que respondem em tempo real às requisições dos usuários (vendas, estoque, cadastro).",
         "Nós de processamento analítico, que consolidam dados e geram relatórios e indicadores de vendas."])
R.h2("1.2 O problema de decisão")
R.p("A empresa precisa tomar duas decisões encadeadas. A primeira é estratégica: escolher o provedor de nuvem, equilibrando "
    "custo, desempenho e confiabilidade. A segunda é operacional: com os preços do provedor escolhido, decidir quantas "
    "instâncias de cada tipo contratar para maximizar a capacidade de atendimento sem ultrapassar a cota de vCPU, a memória "
    "disponível e o orçamento mensal.")
R.caixa("Como as etapas se conectam", [
    "Etapa 1 (AHP) → define o provedor e, com ele, o preço por hora de cada tipo de instância.",
    "Etapa 2 (PL) → usa esses preços como coeficientes da restrição de orçamento e encontra a alocação ótima.",
])

# ---- 2. etapa 1
R.h1("2. Etapa 1 — Estudo Multicritério (AHP e Normalização)")
R.h2("2.1 Alternativas pesquisadas")
R.p("Para comparar alternativas equivalentes, fixou-se uma instância de 2 vCPU e 8 GB de RAM, Linux, sob demanda, na região "
    "de São Paulo sempre que o provedor a oferece (a DigitalOcean não tem região no Brasil). O custo mensal considera 730 "
    "horas de uso e câmbio de R$ 5,4701 por dólar (cotação de compra da AwesomeAPI em 04/11/2026, 14h18).")
R.tabela("Alternativas e preços (instância 2 vCPU / 8 GB, sob demanda)",
         ["Provedor", "Instância", "Região", "US$/hora", "R$/mês"],
         [[i, r["Instância"], r["Região"], f"{r['USD_hora']:.4f}".replace(".", ","), brl(r["Custo mensal (R$)"])]
          for i, r in dados.iterrows()],
         larguras=[2.9, 4.6, 4.0, 1.7, 2.6],
         fonte="APIs e páginas de preço de cada provedor (links na Tabela 2). A DigitalOcean cobra no máximo "
               "US$ 63/mês; o valor por hora mostrado é o equivalente em 730 h.")
R.tabela("Fontes consultadas",
         ["Provedor", "Preço", "SLA"],
         [[i, r["Fonte_preco"], r["Fonte_SLA"]] for i, r in dados.iterrows()],
         larguras=[2.9, 7.5, 5.4], tamanho=7.5, alinhar_num=False)

R.h2("2.2 Critérios de avaliação")
R.tabela("Critérios, classificação e justificativa",
         ["Critério", "Tipo", "Unidade", "Justificativa técnica"],
         [["C1 Custo mensal", "Custo ↓", "R$/mês", "Gasto recorrente (OPEX); define a viabilidade financeira da startup."],
          ["C2 Latência", "Custo ↓", "ms", "Tempo de conexão até o datacenter; afeta diretamente a experiência do usuário."],
          ["C3 SLA", "Benefício ↑", "%", "Disponibilidade garantida em contrato para uma VM única; reduz o risco de parada."],
          ["C4 IOPS", "Benefício ↑", "op/s", "Desempenho do disco padrão; limita bancos de dados e filas."],
          ["C5 Suporte", "Benefício ↑", "nota 1–5", "Qualidade do suporte técnico, medida por uma régua própria."]],
         larguras=[3.0, 2.2, 2.0, 8.6])
R.p("A nota de suporte (1 a 5) foi atribuída pelo autor considerando cinco aspectos dos planos de suporte de cada "
    "provedor: atendimento 24/7, suporte técnico incluído sem custo adicional, atendimento em português, canais de chat "
    "ou telefone e tempo de resposta garantido em contrato.")
R.tabela("Matriz de decisão com os valores brutos",
         ["Provedor"] + CRITERIOS,
         [[i] + [brl(r[c]) if c.startswith("Custo") else num(r[c]) for c in CRITERIOS] for i, r in dados.iterrows()],
         fonte="Latência: mediana de 10 conexões TCP na porta 443 medidas de Araraquara em 04/11/2026.")
R.caixa("Como a latência foi medida", [
    "Com o comando curl (tempo de conexão TCP na porta 443), 10 vezes por provedor, a partir de um computador em "
    "Araraquara-SP. Usou-se TCP em vez de ping porque AWS e Azure bloqueiam ICMP. Endpoints: ec2.sa-east-1.amazonaws.com "
    "(AWS), s3brazilsouth.blob.core.windows.net (Azure), endpoint gcping de southamerica-east1 (Google), "
    "objectstorage.sa-saopaulo-1.oraclecloud.com (Oracle), nyc3.digitaloceanspaces.com (DigitalOcean) e "
    "br-gru-1.linodeobjects.com (Linode).",
])

R.h2("2.3 Normalização Min-Max")
R.caixa("Fórmulas aplicadas", [
    "Benefício (quanto maior, melhor):   Nij = (Xij − Xj,mín) / (Xj,máx − Xj,mín)",
    "Custo (quanto menor, melhor):   Nij = (Xj,máx − Xij) / (Xj,máx − Xj,mín)",
])
R.tabela("Mínimos, máximos e amplitude de cada critério",
         ["Critério", "Tipo", "Mínimo", "Máximo", "Amplitude"],
         [[c, "Custo" if TIPOS[c] == "custo" else "Benefício", num(dados[c].min()), num(dados[c].max()),
           num(round(dados[c].max() - dados[c].min(), 4))] for c in CRITERIOS])
ex = "Google Cloud" if "Google Cloud" in dados.index else rk.index[2]
c0, c2 = CRITERIOS[0], CRITERIOS[2]
R.p("Exemplos de memória de cálculo:", negrito=True)
R.lista([f"{ex}, {c0} (custo): N = ({n2(dados[c0].max())} − {n2(dados.loc[ex, c0])}) / ({n2(dados[c0].max())} − "
         f"{n2(dados[c0].min())}) = {n4(norm.loc[ex, c0])}",
         f"{ex}, {c2} (benefício): N = ({num(dados.loc[ex, c2])} − {num(dados[c2].min())}) / ({num(dados[c2].max())} − "
         f"{num(dados[c2].min())}) = {n4(norm.loc[ex, c2])}"])
R.tabela("Matriz normalizada Nij", ["Provedor"] + CURTOS,
         [[i] + [n4(x) for x in r] for i, r in norm.iterrows()],
         destaques={(li, ci + 1): VERDE for li, (_, r) in enumerate(norm.iterrows()) for ci, x in enumerate(r) if x == 1.0})
R.p("Células em verde indicam o melhor desempenho em cada critério (N = 1).", italico=True, tamanho=9, alinhar="esq")

R.h2("2.4 Pesos dos critérios pelo AHP de Saaty")
R.p("Julgamentos do autor: para uma startup em fase inicial, o custo é o critério mais importante (moderadamente mais "
    "importante que latência e SLA; fortemente mais importante que IOPS e suporte). Latência e SLA têm a mesma importância "
    "entre si e são moderadamente mais importantes que IOPS e suporte, que empatam.")
R.tabela("Matriz de comparação par a par A (Escala Fundamental de Saaty)",
         [""] + CURTOS,
         [[CURTOS[i]] + [fmt(__import__("fractions").Fraction(M[i][j]).limit_denominator(9)) for j in range(5)] for i in range(5)]
         + [["Soma"] + [n4(x) for x in ahp["soma_colunas"]]],
         destaques={(5, None): AZUL_CLARO})
R.tabela("Matriz normalizada (cada coluna dividida pela sua soma) e vetor de pesos",
         [""] + CURTOS + ["Peso W"],
         [[CURTOS[i]] + [n4(x) for x in ahp["normalizada"][i]] + [n4(ahp["pesos"][i])] for i in range(5)],
         destaques={(i, 6): AZUL_CLARO for i in range(5)})
R.figura("Peso de cada critério", g_pesos, largura=14)
R.tabela("Verificação de consistência", ["Critério", "(A·w)i", "wi", "(A·w)i / wi"],
         [[CURTOS[i], n4(ahp["aw"][i]), n4(ahp["pesos"][i]), n4(ahp["aw"][i] / ahp["pesos"][i])] for i in range(5)])
R.caixa("Resultado da consistência", [
    f"λmax = {n4(ahp['lambda_max'])}",
    f"IC = (λmax − n) / (n − 1) = ({n4(ahp['lambda_max'])} − 5) / 4 = {n4(ahp['ic'])}",
    f"RC = IC / IR = {n4(ahp['ic'])} / 1,12 = {n4(ahp['rc'])}",
    "Como RC < 0,10, os julgamentos são consistentes." if ahp["consistente"] else "RC ≥ 0,10: revisar julgamentos.",
], cor=VERDE, borda="2E7D4F")

R.h2("2.5 Ranking final")
R.tabela("Pontuação global = Σ Wj · Nij", ["Posição", "Provedor"] + CURTOS + ["Pontuação"],
         [[f"{int(r['Posição'])}º", i] + [n4(contrib.loc[i, c]) for c in CRITERIOS] + [n4(r["Pontuação"])]
          for i, r in rk.iterrows()],
         destaques={(0, None): VERDE}, primeira_negrito=False)
R.figura("Composição da pontuação de cada provedor por critério", g_ranking, largura=15)
R.h2("2.6 Parecer técnico")
lv, ls = rk.loc[vencedor, "Pontuação"], rk.loc[segundo, "Pontuação"]
R.caixa(f"Recomendação: {vencedor}", [
    f"{vencedor} obteve a maior pontuação ({n4(lv)}), {n4(lv - ls)} acima de {segundo}. A vantagem vem principalmente do "
    f"critério de maior peso: o custo mensal de {brl(dados.loc[vencedor, c0])} é o menor entre os seis provedores, cerca de "
    f"{(1 - dados.loc[vencedor, c0] / dados[c0].max()) * 100:.0f}% mais barato que a opção mais cara.",
    f"Os pontos de atenção são o suporte (nota {num(dados.loc[vencedor, 'Suporte (1-5)'])}) e o IOPS do disco padrão, "
    f"em que {segundo} é superior. A latência é praticamente igual entre os provedores com região em São Paulo "
    f"(31 a 35 ms); apenas a DigitalOcean, sem região no Brasil, fica muito atrás.",
], cor=VERDE, borda="2E7D4F")

# ---- 3. etapa 2
R.h1("3. Etapa 2 — Programação Linear e Método Simplex")
R.h2("3.1 Formulação do modelo")
R.p(f"Com {vencedor} escolhido, a empresa dimensiona a infraestrutura de produção. As variáveis de decisão são:")
R.tabela("Variáveis de decisão", ["Variável", "Significado", "Configuração", "Capacidade", "Custo"],
         [["X1", "Instâncias de microsserviços transacionais", "2 vCPU · 8 GB", "500 req/s", "R$ 148/mês"],
          ["X2", "Nós de processamento analítico", "4 vCPU · 32 GB", "1.200 req/s", "R$ 391/mês"]],
         larguras=[1.6, 5.6, 2.9, 2.6, 2.6])
R.p("Os custos vêm da tabela oficial da Oracle (US$ 0,025 por OCPU-hora e US$ 0,0015 por GB-hora; 1 OCPU = 2 vCPU): "
    "X1 = 1 OCPU + 8 GB = US$ 0,037/h ≈ R$ 148/mês; X2 = 2 OCPU + 32 GB = US$ 0,098/h ≈ R$ 391/mês.")
R.p("As capacidades de 500 e 1.200 req/s são premissas de projeto estimadas pelo autor. O nó analítico tem o dobro de "
    "vCPU e quatro vezes a memória do microsserviço, mas processa consultas mais pesadas, por isso sua capacidade "
    "equivalente foi estimada em 2,4 vezes a do microsserviço, e não em 4 vezes. Em produção, esses valores seriam "
    "calibrados com um teste de carga (por exemplo, com as ferramentas k6 ou JMeter).")
R.caixa("Modelo matemático", [
    f"Max Z = {C[0]}X1 + {num(C[1])}X2   (requisições por segundo)",
    "sujeito a:",
    f"   {A[0][0]}X1 + {A[0][1]}X2 ≤ {B[0]}   (cota de vCPU da conta)",
    f"   {A[1][0]}X1 + {A[1][1]}X2 ≤ {B[1]}   (memória contratada, GB)",
    f"   {A[2][0]}X1 + {A[2][1]}X2 ≤ {num(B[2])}   (orçamento mensal, R$)",
    "   X1, X2 ≥ 0",
], mono=True)
R.p("Os limites são premissas do cenário: 64 vCPU é a cota de serviço da conta da Cloud Search no provedor, 320 GB é a "
    "memória total contratada para o ambiente de produção e R$ 6.000 é a verba mensal de TI definida pela startup.")

R.h2("3.2 Forma padrão")
R.caixa("Forma padrão com variáveis de folga", [
    f"Max Z = {C[0]}X1 + {num(C[1])}X2 + 0X3 + 0X4 + 0X5",
    f"   {A[0][0]}X1 + {A[0][1]}X2 + X3 = {B[0]}",
    f"   {A[1][0]}X1 + {A[1][1]}X2 + X4 = {B[1]}",
    f"   {A[2][0]}X1 + {A[2][1]}X2 + X5 = {num(B[2])}",
    "   X1, X2, X3, X4, X5 ≥ 0",
], mono=True)
R.tabela("Significado físico das variáveis de folga", ["Folga", "Recurso", "Significado no sistema de TI"],
         [[f"X{3+i}", RECURSOS[i], s] for i, s in enumerate(SIGNIF_FOLGA)], larguras=[1.6, 4.0, 10.2])
R.p("As colunas de X3, X4 e X5 formam a Matriz Identidade inicial (I3), por isso as folgas compõem a base do primeiro quadro "
    "e a solução inicial é X1 = X2 = 0, Z = 0.")

R.h2("3.3 Iterações do Simplex Tabular")
R.p("Legenda das cores: amarelo = coluna pivô (variável que entra); azul = linha pivô (variável que sai); "
    "laranja = elemento pivô.", italico=True, tamanho=9, alinhar="esq")
for passo in sol["passos"]:
    nome = ("Quadro inicial — Acampamento Base (Z = 0)" if passo["iteracao"] == 0
            else f"Quadro após a iteração {passo['iteracao']}")
    if passo["status"] == "otimo":
        nome += " — solução ótima"
    R.h3(nome)
    tem_razao = passo["status"] == "continua"
    cab = ["Base", "Cb"] + passo["variaveis"] + ["RHS"] + (["Razão"] if tem_razao else [])
    linhas = []
    for i, bv in enumerate(passo["base"]):
        lin = [bv, fmt(passo["cb"][i])] + [fmt(x) for x in passo["tabela"][i]]
        if tem_razao:
            rz = passo["razoes"][i]
            lin.append(fmt(rz) if rz is not None else "—")
        linhas.append(lin)
    linhas.append(["Zj", ""] + [fmt(x) for x in passo["zj"]] + [fmt(passo["z"])] + ([""] if tem_razao else []))
    linhas.append(["Cj − Zj", ""] + [fmt(x) for x in passo["cj_zj"]] + [""] + ([""] if tem_razao else []))
    dest = {(len(linhas) - 1, None): AZUL_CLARO, (len(linhas) - 2, None): CINZA}
    if tem_razao:
        col = 2 + passo["col_pivo"]
        for li in range(len(linhas)):
            dest[(li, col)] = AMARELO
        for ci in range(len(cab)):
            dest[(passo["lin_pivo"], ci)] = AZUL_LINHA
        dest[(passo["lin_pivo"], col)] = LARANJA
    R.tabela(nome, cab, linhas, destaques=dest)
    if tem_razao:
        R.lista([f"Variável que entra: {passo['entra']}, pois tem o maior Cj − Zj positivo ({fmt(passo['cj_zj'][passo['col_pivo']])}).",
                 f"Variável que sai: {passo['sai']}, pois tem a menor razão RHS/coluna pivô "
                 f"({fmt(passo['razoes'][passo['lin_pivo']])}); é o recurso que se esgota primeiro.",
                 f"Elemento pivô: {fmt(passo['pivo'])}."])
        R.p("Operações de Gauss-Jordan:", negrito=True, alinhar="esq", depois=2)
        R.caixa(None, passo["operacoes"], cor=CINZA, borda="7FA7D1", mono=True)
    else:
        R.caixa(None, ["Critério de parada atingido: todos os Cj − Zj ≤ 0. A solução é ótima."], cor=VERDE, borda="2E7D4F")

R.h2("3.4 Leitura gerencial (Business Decision Dashboard)")
R.caixa("Decisão ótima", [
    f"Contratar {fmt(v['X1'])} instâncias de microsserviços e {fmt(v['X2'])} nós analíticos.",
    f"Capacidade máxima: Z = {num(int(sol['z']))} requisições por segundo.",
    f"Custo mensal: {brl(float(B[2] - v['X5']))} de {brl(B[2])}.",
], cor=VERDE, borda="2E7D4F")
linhas = []
dest = {}
for i, rec in enumerate(RECURSOS):
    folga = v[f"X{3+i}"]
    gargalo = folga == 0
    linhas.append([rec, num(B[i]), num(int(B[i] - folga)), num(int(folga)), "Gargalo" if gargalo else "Capacidade ociosa"])
    dest[(i, 4)] = LARANJA if gargalo else VERDE
R.tabela("Uso dos recursos na solução ótima", ["Recurso", "Disponível", "Usado", "Folga", "Situação"], linhas, destaques=dest)
R.figura("Percentual usado de cada recurso", g_recursos, largura=14)
R.p("vCPU e memória são os gargalos operacionais (folga zero): para aumentar a capacidade, a empresa precisa ampliar a "
    f"cota de vCPU e a memória, e não o orçamento. Sobram {brl(float(v['X5']))} por mês, que podem ser destinados a backup, "
    "monitoramento ou a uma reserva para picos de demanda.")

# ---- 4. dashboard
R.h1("4. Dashboard Utilitário")
R.h2("4.1 Arquitetura")
R.p("A ferramenta foi desenvolvida em Python com Streamlit (interface), Pandas e NumPy (cálculos) e Plotly (gráficos). "
    "O motor do Simplex foi implementado em Python, com apoio de IA documentado no Apêndice, usando frações exatas "
    "(módulo fractions), o que elimina erros de arredondamento no pivoteamento. Cada resultado é validado automaticamente com scipy.optimize.linprog.")
R.tabela("Organização do código", ["Arquivo", "Responsabilidade"], [
    ["app.py", "Interface com as abas Dados, AHP e Simplex"],
    ["motores/ahp.py", "Normalização Min-Max, pesos de Saaty, λmax, IC, RC e ranking"],
    ["motores/simplex.py", "Simplex Tabular com registro de cada quadro e das operações de Gauss-Jordan"],
    ["motores/precos.py", "Consulta de preço ao vivo na Azure Retail Prices API"],
    ["dados/pesquisa_mercado.csv", "Pesquisa de mercado com fontes e status de verificação"],
    ["gerar_relatorio.py", "Gera este relatório a partir dos mesmos motores"],
], larguras=[4.8, 11.0], alinhar_num=False)
R.h2("4.2 Como usar")
R.lista(["Instalar as dependências: pip install -r requirements.txt",
         "Iniciar o app: streamlit run app.py",
         "Aba Dados: editar a pesquisa, ajustar o câmbio e buscar o preço da Azure ao vivo.",
         "Aba AHP: escolher os julgamentos de Saaty e acompanhar pesos, RC e ranking.",
         "Aba Simplex: editar coeficientes e limites e acompanhar cada quadro até a solução ótima.",
         "Repositório: https://github.com/Felpz016/Projeto_PL"], numerada=True)
R.h2("4.3 Telas")


def partes_do_print(caminho, max_ratio=1.25):
    """Recorta a barra lateral e o espaço em branco do print e divide prints altos em partes."""
    from PIL import Image
    import numpy as np
    im = Image.open(caminho).convert("RGB")
    w, h = im.size
    im = im.crop((int(w * 300 / 1400), 0, w, h))  # remove a barra lateral do Streamlit
    arr = np.asarray(im).astype(int)
    branca = (arr.min(axis=2) > 245).all(axis=1)
    ult = len(branca) - 1
    while ult > 0 and branca[ult]:
        ult -= 1
    im = im.crop((0, 0, im.size[0], min(ult + 30, im.size[1])))
    branca = branca[: im.size[1]]
    w, h = im.size
    n = max(1, int(np.ceil(h / (w * max_ratio))))
    cortes, inicio = [0], 0
    for k in range(1, n):
        alvo = int(h * k / n)
        candidatos = [y for y in range(max(inicio + 50, alvo - 250), min(h - 50, alvo + 250)) if branca[y]]
        y = min(candidatos, key=lambda c: abs(c - alvo)) if candidatos else alvo
        cortes.append(y)
        inicio = y
    cortes.append(h)
    saidas = []
    for k in range(len(cortes) - 1):
        buf = io.BytesIO()
        im.crop((0, cortes[k], w, cortes[k + 1])).save(buf, format="PNG")
        buf.seek(0)
        saidas.append(buf)
    return saidas


for i, legenda in enumerate(["Aba Dados de mercado", "Aba AHP e Ranking", "Aba Simplex e painel gerencial"], 1):
    img = BASE / "prints" / f"aba{i}.png"
    if img.exists():
        partes = partes_do_print(img)
        for k, parte in enumerate(partes, 1):
            sufixo = f" (parte {k} de {len(partes)})" if len(partes) > 1 else ""
            R.figura(legenda + sufixo, parte, largura=15.0)

# ---- 5. apêndice A: IA
def bloco_codigo(caminho, titulo):
    """Código-fonte completo em uma caixa cinza, fonte monoespaçada."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    bordas_tabela(t, "D5DBE3", 4)
    margens_celula(t, 0.2)
    cel = t.rows[0].cells[0]
    sombrear(cel, "F7F8FA")
    pp = cel.paragraphs[0]
    rr = pp.add_run(titulo)
    rr.bold = True
    rr.font.size = Pt(8)
    rr.font.color.rgb = AZUL
    for linha in (BASE / caminho).read_text(encoding="utf-8").rstrip().splitlines():
        lp = cel.add_paragraph()
        lp.paragraph_format.space_after = Pt(0)
        lp.paragraph_format.line_spacing = 1.0
        lr = lp.add_run(linha.replace("\t", "    ") or " ")
        lr.font.name = "Consolas"
        lr.font.size = Pt(7)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def ficha(numero, etapa, objetivo, prompt, entrega, auditoria, arquivos):
    R.h3(f"Ficha {numero} — {etapa}")
    linhas = [["1. Ferramenta", "Claude Opus 5.5 (Anthropic), usado pelo Claude Code"],
              ["2. Etapa vinculada", etapa],
              ["3. Objetivo técnico", objetivo],
              ["4. Prompt integral", prompt],
              ["5. Solução entregue pela IA", entrega],
              ["6. Auditoria humana e teste de mesa", auditoria]]
    t = doc.add_table(rows=0, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    bordas_tabela(t)
    margens_celula(t, 0.12)
    t.autofit = False
    for k, val in linhas:
        c0, c1 = t.add_row().cells
        c0.width, c1.width = Cm(4.2), Cm(11.6)
        sombrear(c0, AZUL_CLARO)
        r0 = c0.paragraphs[0].add_run(k)
        r0.bold = True
        r0.font.size = Pt(8.5)
        r0.font.color.rgb = AZUL
        c1.paragraphs[0].paragraph_format.space_after = Pt(0)
        texto_com_pendencias(c1.paragraphs[0], val, tamanho=8.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    for caminho, titulo in arquivos:
        bloco_codigo(caminho, titulo)


P01 = "de acordo com o arquivo da proposta do projeto, me de ideias de como realiza-lo"
P02 = "pode montar pra eu ver se é isso mesmo que quero"
P03 = ("faça oque for possivel no momento, estou em outro computador, quando voltar pra casa voce faz o resto que nao foi "
       "possivel” e, depois, “pode continuar a fazer tudo que faltava, mas agora nesse computador")
P04 = "refaça o relatorio do projeto de um jeito mais bonito, e me mande as perguntas que preciso completar"
P05 = "crie um repositorio para mim: https://github.com/Felpz016 (meu perfil)"

R.h1("5. Apêndice A — Governança e Rastreabilidade do Uso de IA")
R.p("A IA generativa foi usada como copiloto em cinco momentos: planejamento, código do dashboard, pesquisa de dados, "
    "redação do relatório e publicação do código no GitHub. Esta seção segue o modelo da Seção 6 da proposta: matriz de rastreabilidade, correções feitas "
    "durante a validação e uma ficha detalhada para cada script gerado com apoio de IA.")
R.h2("5.1 Matriz de rastreabilidade")
R.tabela("Matriz de Rastreabilidade de IA",
         ["ID / Etapa", "IA e versão", "Objetivo", "Prompt (entrada exata)", "Saída gerada", "Auditoria e validação"], [
    ["#01 Planejamento", "Claude Opus 5.5 (Anthropic)", "Sugerir uma abordagem para o projeto",
     f"“{P01}”", "Sugestão do Cenário B (Cloud), dos 5 critérios e do modelo de PL",
     "O autor escolheu o Cenário B e aprovou a abordagem."],
    ["#02 Etapas 1 e 2", "Claude Opus 5.5 (Anthropic)", "Gerar o dashboard e os motores de cálculo",
     f"“{P02}”", "app.py, motores/ahp.py e motores/simplex.py (Fichas 1 a 3)",
     "Simplex comparado com o SciPy (resultados iguais). Teste de mesa refeito pelo autor (Apêndice B): os "
     "resultados conferem."],
    ["#03 Etapa 1", "Claude Opus 5.5 (Anthropic)", "Levantar preços, SLA e IOPS e medir a latência",
     f"“{P03}”", "dados/pesquisa_mercado.csv com as fontes da Tabela 2",
     "Latência medida no computador do autor; preços e SLAs conferidos nas fontes oficiais (ver 5.2)."],
    ["#04 Documentação", "Claude Opus 5.5 (Anthropic)", "Gerar e formatar este relatório",
     f"“{P04}”", "gerar_relatorio.py e este documento",
     "O autor forneceu nome, RA e nome da startup e aprovou as premissas sugeridas (notas de suporte, "
     "julgamentos de Saaty, capacidades e limites do PL)."],
    ["#05 Publicação", "Claude Opus 5.5 (Anthropic)", "Enviar o código para o GitHub",
     f"“{P05}”", "Commit do código no repositório github.com/Felpz016/Projeto_PL",
     "O autor criou o repositório na própria conta do GitHub; o envio usou o login dele."],
], tamanho=7.5, alinhar_num=False, larguras=[2.0, 2.2, 2.4, 3.8, 2.7, 2.7])

R.h2("5.2 Correções feitas durante a validação")
R.p("A validação encontrou e corrigiu erros e imprecisões nas primeiras versões geradas pela IA:")
R.lista([
    "SLA da Oracle: a primeira versão usava 99,95%, valor informado de memória pela IA. A conferência no documento "
    "oficial da Oracle (Pillar Document, set/2026, seção 3.6.21) mostrou 99,9% para instância única, e o valor foi corrigido.",
    "Latência: a primeira versão usava valores estimados. Eles foram substituídos pela mediana de 10 medições reais "
    "feitas a partir de Araraquara.",
    "Preço da Oracle: a primeira versão usava preços por OCPU e por GB informados de memória. Eles foram confirmados na "
    "API oficial de preços da Oracle (itens B93113 e B93114).",
    "Fontes: a fonte do SLA da Azure apontava para o site da Azure na China e foi trocada pelo SLA oficial global da "
    "Microsoft. Os preços da AWS e do Google vinham de sites de terceiros e foram confirmados nas páginas oficiais "
    "(AWS: US$ 0,1530/h; Google: US$ 0,1063704/h). A região da DigitalOcean foi alinhada para NYC3, a mesma em que a "
    "latência foi medida.",
    "Dashboard: um texto com “R$” era interpretado como fórmula matemática pelo Streamlit e aparecia quebrado. O símbolo "
    "foi escapado no código.",
])

R.h2("5.3 Fichas detalhadas dos scripts")
ficha(1, "Etapa 1 (AHP de Saaty e Normalização)",
      "Normalizar a matriz de decisão 6×5 distinguindo critérios de custo e de benefício, calcular os pesos de Saaty, "
      "λmax, IC e RC e gerar o ranking por ponderação aditiva.",
      f"“{P02}” (o código foi gerado junto com o protótipo do dashboard).",
      "Módulo motores/ahp.py, com as funções normalizar_min_max, matriz_de_triangulo, calcular_ahp e ranking "
      "(código completo abaixo).",
      "Os resultados das funções foram conferidos com a memória de cálculo das Tabelas 5 a 9. O teste de mesa da "
      "normalização (Google Cloud, critério Custo) e do peso do critério Custo foi refeito à mão pelo autor "
      "(Apêndice B) e confere.",
      [("motores/ahp.py", "motores/ahp.py — código completo")])
ficha(2, "Etapa 2 (Programação Linear e Simplex Tabular)",
      "Executar o Simplex Tabular para maximização com frações exatas, registrando cada quadro, o Cj − Zj, o teste da "
      "razão mínima, o pivô e as operações de Gauss-Jordan, com critério de parada Cj − Zj ≤ 0.",
      f"“{P02}” (o código foi gerado junto com o protótipo do dashboard).",
      "Módulo motores/simplex.py, com a função resolver (código completo abaixo).",
      "O resultado (X1 = 24, X2 = 4, Z = 16.800) é idêntico ao do scipy.optimize.linprog. O teste de mesa do 1º "
      "pivoteamento e a prova real da solução ótima foram refeitos à mão pelo autor (Apêndice B) e conferem.",
      [("motores/simplex.py", "motores/simplex.py — código completo")])
ficha(3, "Interface do Dashboard",
      "Construir a interface interativa com as abas Dados, AHP e Simplex, a busca de preço ao vivo da Azure e o painel "
      "de gargalos e folgas.",
      f"“{P02}”",
      "Arquivos app.py (interface Streamlit) e motores/precos.py (API pública de preços da Azure). Pelo tamanho, o "
      "código completo do app.py está no repositório; abaixo está o módulo de preços.",
      "O app foi executado e conferido tela a tela (Seção 4.3). A busca ao vivo retornou US$ 0,153/h para o "
      "D2s v5 em Brazil South, igual ao valor da Tabela 1.",
      [("motores/precos.py", "motores/precos.py — código completo")])

R.h2("5.4 Declaração de uso de IA")
R.p("Declaro que as ferramentas de Inteligência Artificial listadas neste apêndice foram usadas como copiloto no "
    "planejamento, na programação do dashboard, na pesquisa de dados e na redação deste relatório, e que todas as "
    "interações relevantes estão documentadas acima.")
doc.add_paragraph()
assin = doc.add_paragraph("_____________________________________________")
assin.alignment = WD_ALIGN_PARAGRAPH.CENTER
assin.paragraph_format.space_after = Pt(0)
R.p("Felipe Benaglia Nardini — RA 2880482421038", alinhar="centro")

# ---- 6. apêndice B: teste de mesa
R.h1("6. Apêndice B — Teste de Mesa (Conferência Manual)")
R.p("Este roteiro refaz à mão, com uma calculadora comum, quatro contas que o dashboard faz automaticamente. Se os "
    "resultados baterem com os valores esperados, os scripts gerados com apoio de IA estão corretos nesses pontos.")
R.caixa(None, ["Conferido pelo autor em 04/11/2026: os quatro resultados refeitos na calculadora batem com os "
               "valores esperados."], cor=VERDE, borda="2E7D4F")

gc = "Google Cloud"
cmax, cmin, cg = dados[c0].max(), dados[c0].min(), dados.loc[gc, c0]
R.h2("B.1 Normalização do Google Cloud no critério Custo")
R.caixa(None, [
    f"Dados: maior custo = {n2(cmax)} (AWS e Azure); menor custo = {n2(cmin)} (Oracle); Google = {n2(cg)}.",
    "Critério de custo:  N = (máx − valor) / (máx − mín)",
    f"N = ({n2(cmax)} − {n2(cg)}) / ({n2(cmax)} − {n2(cmin)}) = {n2(cmax - cg)} / {n2(cmax - cmin)} = {n4(norm.loc[gc, c0])}",
    f"Resultado esperado: {n4(norm.loc[gc, c0])} (Tabela 6).   ☑ Conferido",
], mono=True)

R.h2("B.2 Peso do critério Custo pelo AHP")
somas = ahp["soma_colunas"]
linha_norm = [M[0][j] / somas[j] for j in range(5)]
R.caixa(None, [
    "Passo 1: somar cada coluna da matriz de Saaty (Tabela 7):",
    "  Custo: 1 + 1/3 + 1/3 + 1/5 + 1/5 = " + n4(somas[0]) + "   ·   Latência e SLA: 3 + 1 + 1 + 1/3 + 1/3 = " + n4(somas[1]),
    "  IOPS e Suporte: 5 + 3 + 3 + 1 + 1 = " + n4(somas[3]),
    "Passo 2: dividir cada valor da linha Custo pela soma da sua coluna:",
    "  " + "   ".join(f"{fmt(__import__('fractions').Fraction(M[0][j]).limit_denominator(9))}/{n4(somas[j])} = {n4(linha_norm[j])}" for j in range(5)),
    "Passo 3: tirar a média da linha:",
    "  W(Custo) = (" + " + ".join(n4(x) for x in linha_norm) + f") / 5 = {n4(sum(linha_norm))} / 5 = {n4(ahp['pesos'][0])}",
    f"Resultado esperado: {n4(ahp['pesos'][0])} (Tabela 8).   ☑ Conferido",
], mono=True)

p0 = sol["passos"][0]
p1 = sol["passos"][1]
R.h2("B.3 Primeiro pivoteamento do Simplex")
R.caixa(None, [
    f"Entra X2: maior Cj − Zj do quadro inicial ({fmt(p0['cj_zj'][1])}).",
    f"Razões: 64 ÷ 4 = 16;  320 ÷ 32 = 10;  6.000 ÷ 391 ≈ 15,35.  Menor = 10 → sai X4 (RAM). Pivô = 32.",
    "Nova L2 = L2 ÷ 32 = [8, 32, 0, 1, 0 | 320] ÷ 32 = [1/4, 1, 0, 1/32, 0 | 10]",
    "Nova L1 = L1 − 4 × Nova L2 = [2 − 1, 4 − 4, 1, 0 − 1/8, 0 | 64 − 40] = [1, 0, 1, −1/8, 0 | 24]",
    "Nova L3 = L3 − 391 × Nova L2 = [148 − 391/4, 0, 0, −391/32, 1 | 6.000 − 3.910] = [201/4, 0, 0, −391/32, 1 | 2.090]",
    f"Z = 1.200 × 10 = {num(int(p1['z']))}.   Novo Cj − Zj de X1 = 500 − 1.200 × 1/4 = {fmt(p1['cj_zj'][0])}.",
    "Resultado esperado: igual ao quadro após a iteração 1 (Tabela 14).   ☑ Conferido",
], mono=True)

R.h2("B.4 Prova real da solução ótima")
x1, x2 = int(v["X1"]), int(v["X2"])
R.caixa(None, [
    f"Substituindo X1 = {x1} e X2 = {x2} nas restrições:",
    f"  vCPU:      2 × {x1} + 4 × {x2} = {2*x1} + {4*x2} = {2*x1 + 4*x2} ≤ 64   → folga 0 (gargalo)",
    f"  RAM:       8 × {x1} + 32 × {x2} = {8*x1} + {32*x2} = {8*x1 + 32*x2} ≤ 320   → folga 0 (gargalo)",
    f"  Orçamento: 148 × {x1} + 391 × {x2} = {num(148*x1)} + {num(391*x2)} = {num(148*x1 + 391*x2)} ≤ 6.000   → folga {num(6000 - 148*x1 - 391*x2)}",
    f"  Z = 500 × {x1} + 1.200 × {x2} = {num(500*x1)} + {num(1200*x2)} = {num(500*x1 + 1200*x2)} req/s",
    f"Resultado esperado: Z = {num(int(sol['z']))} e folgas 0, 0 e {num(int(v['X5']))} (Tabela 16).   ☑ Conferido",
], mono=True)

saida = BASE / "relatorio" / "Relatorio_Projeto_Final_PO.docx"
saida.parent.mkdir(exist_ok=True)
doc.save(saida)
print("Relatório gerado em", saida)

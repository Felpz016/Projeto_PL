"""Dashboard de Apoio à Decisão: AHP (Etapa 1) + Simplex Tabular (Etapa 2).

Rodar:  pip install -r requirements.txt  &&  streamlit run app.py
"""
from fractions import Fraction

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from motores.ahp import calcular_ahp, matriz_de_triangulo, normalizar_min_max, ranking
from motores.simplex import fmt, resolver
from motores import precos

st.set_page_config(page_title="Apoio à Decisão Cloud", page_icon="☁️", layout="wide")

# ---------------------------------------------------------------- dados iniciais
# Pesquisa de mercado: instância equivalente 2 vCPU / 8 GB, Linux, sob demanda.
# A coluna Status diz o que já foi conferido e o que ainda precisa ser confirmado.
from pathlib import Path

CSV_PESQUISA = Path(__file__).parent / "dados" / "pesquisa_mercado.csv"
HORAS_MES = 730


def carregar_pesquisa(cambio: float) -> pd.DataFrame:
    df = pd.read_csv(CSV_PESQUISA)
    df.insert(df.columns.get_loc("USD_hora") + 1, "Custo mensal (R$)", (df["USD_hora"] * HORAS_MES * cambio).round(2))
    return df


CRITERIOS = ["Custo mensal (R$)", "Latência SP (ms)", "SLA (%)", "IOPS", "Suporte (1-5)"]
TIPOS_PADRAO = {"Custo mensal (R$)": "custo", "Latência SP (ms)": "custo",
                "SLA (%)": "beneficio", "IOPS": "beneficio", "Suporte (1-5)": "beneficio"}

ESCALA = {"1/9": 1/9, "1/8": 1/8, "1/7": 1/7, "1/6": 1/6, "1/5": 1/5, "1/4": 1/4, "1/3": 1/3, "1/2": 1/2,
          "1": 1, "2": 2, "3": 3, "4": 4, "5": 5, "6": 6, "7": 7, "8": 8, "9": 9}
# julgamentos iniciais (linha i vs coluna j, i < j)
JULG_PADRAO = {(0, 1): "3", (0, 2): "3", (0, 3): "5", (0, 4): "5",
               (1, 2): "1", (1, 3): "3", (1, 4): "3",
               (2, 3): "3", (2, 4): "3",
               (3, 4): "1"}

cambio = st.sidebar.number_input("Câmbio USD → BRL (cotação do dia da pesquisa)", value=5.47, step=0.01, format="%.2f")
st.sidebar.caption("Custo mensal = US$/hora × 730 h × câmbio")
if st.session_state.get("cambio_usado") != cambio:
    st.session_state.dados = carregar_pesquisa(cambio)
    st.session_state.cambio_usado = cambio

st.title("☁️ Sistema de Apoio à Decisão: Provedores Cloud")
st.caption("Etapa 1: AHP de Saaty + Normalização Min-Max · Etapa 2: Programação Linear pelo Simplex Tabular")

aba_dados, aba_ahp, aba_simplex = st.tabs(["1 · Dados de mercado", "2 · AHP e Ranking", "3 · Simplex"])

# ================================================================ ABA 1: DADOS
with aba_dados:
    st.subheader("Matriz de decisão (6 alternativas × 5 critérios)")
    st.info("Preços e SLAs com fonte nas colunas ao lado. Latência = mediana de 10 conexões TCP/443 medidas de "
            "Araraquara em 04/11/2026. A coluna Status diz o que foi conferido.")

    with st.expander("🔄 Buscar preço ao vivo (APIs públicas)"):
        st.write("A Azure Retail Prices API é pública e não pede chave. Os outros provedores ficam com preço manual e link da fonte.")
        col_a, col_b, col_c = st.columns(3)
        sku = col_a.text_input("SKU Azure", "Standard_D2s_v5")
        regiao = col_b.text_input("Região", "brazilsouth")
        col_c.metric("Câmbio em uso", f"{cambio:.2f}")
        if st.button("Buscar Azure"):
            try:
                usd_h = precos.azure_preco_hora(sku, regiao)
                mensal = round(usd_h * 730 * cambio, 2)
                df = st.session_state.dados
                df.loc[df["Provedor"] == "Azure", "Custo mensal (R$)"] = mensal
                df.loc[df["Provedor"] == "Azure", "USD_hora"] = usd_h
                df.loc[df["Provedor"] == "Azure", "Fonte_preco"] = precos.AZURE_URL
                st.success(f"US$ {usd_h:.4f}/h × 730 h × {cambio} = R$ {mensal:,.2f}/mês")
            except Exception as e:  # noqa: BLE001
                st.error(f"Não foi possível buscar: {e}")

    st.session_state.dados = st.data_editor(
        st.session_state.dados, num_rows="dynamic", width='stretch', hide_index=True,
        column_config={"Fonte_preco": st.column_config.LinkColumn("Fonte do preço"),
                       "Fonte_SLA": st.column_config.LinkColumn("Fonte do SLA")},
    )

    st.subheader("Classificação dos critérios")
    cols = st.columns(5)
    tipos = {}
    for k, crit in enumerate(CRITERIOS):
        tipos[crit] = cols[k].radio(crit, ["custo", "beneficio"],
                                    index=0 if TIPOS_PADRAO[crit] == "custo" else 1,
                                    format_func=lambda t: "Custo ↓" if t == "custo" else "Benefício ↑")

dados = st.session_state.dados.set_index("Provedor")

# ================================================================ ABA 2: AHP
with aba_ahp:
    st.subheader("Matriz de comparação par a par (Escala de Saaty)")
    st.caption("Leia como: *o critério da linha é X vezes mais importante que o da coluna*. A parte de baixo é preenchida pela recíproca.")

    julg = {}
    pares = [(i, j) for i in range(5) for j in range(i + 1, 5)]
    grade = st.columns(5)
    for k, (i, j) in enumerate(pares):
        escolha = grade[k % 5].selectbox(f"{CRITERIOS[i].split(' (')[0]} × {CRITERIOS[j].split(' (')[0]}",
                                         list(ESCALA), index=list(ESCALA).index(JULG_PADRAO[(i, j)]), key=f"p{i}{j}")
        julg[(i, j)] = ESCALA[escolha]

    M = matriz_de_triangulo(CRITERIOS, julg)
    res = calcular_ahp(M)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Matriz A**")
        st.dataframe(pd.DataFrame(M, index=CRITERIOS, columns=CRITERIOS).style.format("{:.3f}"))
    with c2:
        st.markdown("**Matriz normalizada (÷ soma da coluna) e pesos W**")
        tab = pd.DataFrame(res["normalizada"], index=CRITERIOS, columns=CRITERIOS)
        tab["Peso W"] = res["pesos"]
        st.dataframe(tab.style.format("{:.4f}"))

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("λmax", f"{res['lambda_max']:.4f}")
    k2.metric("IC = (λmax − n)/(n − 1)", f"{res['ic']:.4f}")
    k3.metric(f"RC = IC / {res['ir']}", f"{res['rc']:.4f}")
    k4.metric("Consistência", "✅ RC < 0,10" if res["consistente"] else "❌ revisar julgamentos")

    st.divider()
    st.subheader("Normalização Min-Max e ranking final")
    norm = normalizar_min_max(dados[CRITERIOS], tipos)
    rk = ranking(norm, res["pesos"])
    c3, c4 = st.columns([3, 2])
    with c3:
        st.markdown("**Matriz normalizada N_ij**")
        st.dataframe(norm.style.format("{:.4f}").background_gradient(cmap="Blues", axis=None))
    with c4:
        st.markdown("**Pontuação = Σ W_j · N_ij**")
        st.dataframe(rk.style.format({"Pontuação": "{:.4f}"}))

    vencedor = rk.index[0]
    st.session_state.vencedor = vencedor
    st.success(f"🏆 Alternativa vencedora: **{vencedor}** ({rk.iloc[0]['Pontuação']:.4f})")

    radar = go.Figure()
    for prov in norm.index:
        radar.add_trace(go.Scatterpolar(r=list(norm.loc[prov]) + [norm.loc[prov].iloc[0]],
                                        theta=CRITERIOS + [CRITERIOS[0]], name=prov, fill="toself", opacity=0.45))
    radar.update_layout(polar=dict(radialaxis=dict(range=[0, 1])), height=460, margin=dict(t=30, b=30))
    st.plotly_chart(radar, width='stretch')

# ================================================================ ABA 3: SIMPLEX
with aba_simplex:
    vencedor = st.session_state.get("vencedor", "—")
    st.subheader(f"Otimização da infraestrutura no provedor vencedor: {vencedor}")
    st.caption("X1 = instâncias de Microsserviços Transacionais (2 vCPU / 8 GB) · X2 = Nós de Processamento Analítico "
               "(4 vCPU / 32 GB) · Max Z = RPS total. Custos padrão: Oracle E4.Flex a R\\$ 5,47/US\\$ (AwesomeAPI, 04/11/2026) "
               "(X1: 1 OCPU + 8 GB ≈ R\\$ 148; X2: 2 OCPU + 32 GB ≈ R\\$ 391). Limites: cota de 64 vCPU, 320 GB de RAM, R\\$ 6.000/mês.")

    preco_venc = float(dados.loc[vencedor, "Custo mensal (R$)"]) if vencedor in dados.index else 150.0
    usar = st.checkbox(f"Usar o preço do vencedor como custo de X1 (R$ {preco_venc:,.2f})", value=False)

    e1, e2 = st.columns(2)
    with e1:
        st.markdown("**Função objetivo (RPS por instância)**")
        c_x1 = st.number_input("c1 (X1)", value=500, step=10)
        c_x2 = st.number_input("c2 (X2)", value=1200, step=10)
    with e2:
        st.markdown("**Coeficientes e capacidades**")
        restr = pd.DataFrame(
            {"Recurso": ["vCPU", "RAM (GB)", "Orçamento (R$/mês)"],
             "X1": [2, 8, round(preco_venc) if usar else 148],
             "X2": [4, 32, round(preco_venc * 0.098 / 0.037) if usar else 391],
             "Limite (b)": [64, 320, 6000]})
        restr = st.data_editor(restr, hide_index=True, width='stretch', key="restr")

    A = restr[["X1", "X2"]].astype(int).values.tolist()
    b = restr["Limite (b)"].astype(int).tolist()
    nomes_folga = [f"X{3+i} (folga {r})" for i, r in enumerate(restr["Recurso"])]

    st.markdown("**Forma padrão**")
    lat = [f"\\text{{Max }} Z = {c_x1}X_1 + {c_x2}X_2 + " + " + ".join(f"0X_{3+i}" for i in range(len(b)))]
    for i, (ai, bi) in enumerate(zip(A, b)):
        lat.append(f"{ai[0]}X_1 + {ai[1]}X_2 + X_{3+i} = {bi}")
    lat.append("X_1, X_2, " + ", ".join(f"X_{3+i}" for i in range(len(b))) + " \\geq 0")
    st.latex(" \\\\ ".join(lat))

    sol = resolver([c_x1, c_x2], A, b, nomes_x=["X1", "X2"], nomes_folga=[f"X{3+i}" for i in range(len(b))])

    for p in sol["passos"]:
        titulo = "Quadro inicial (Acampamento Base, Z = 0)" if p["iteracao"] == 0 else f"Quadro após a iteração {p['iteracao']}"
        if p["status"] == "otimo":
            titulo += " · ÓTIMO"
        with st.expander(titulo, expanded=True):
            cols_v = p["variaveis"]
            linhas = []
            for i, bv in enumerate(p["base"]):
                linhas.append([bv, fmt(p["cb"][i])] + [fmt(v) for v in p["tabela"][i][:-1]] + [fmt(p["tabela"][i][-1])]
                              + ([fmt(p["razoes"][i]) if p.get("razoes") and p["razoes"][i] is not None else "—"]
                                 if p["status"] != "otimo" else []))
            linhas.append(["Zj", ""] + [fmt(v) for v in p["zj"]] + [fmt(p["z"])] + ([""] if p["status"] != "otimo" else []))
            linhas.append(["Cj − Zj", ""] + [fmt(v) for v in p["cj_zj"]] + [""] + ([""] if p["status"] != "otimo" else []))
            colunas = ["Base", "Cb"] + cols_v + ["RHS"] + (["Razão RHS/coluna pivô"] if p["status"] != "otimo" else [])
            df = pd.DataFrame(linhas, columns=colunas)

            def destacar(df_):
                est = pd.DataFrame("", index=df_.index, columns=df_.columns)
                if "col_pivo" in p:
                    est[cols_v[p["col_pivo"]]] = "background-color: rgba(255,200,0,.25)"
                    if "lin_pivo" in p:
                        est.iloc[p["lin_pivo"], :] = "background-color: rgba(0,150,255,.18)"
                        est.iloc[p["lin_pivo"], df_.columns.get_loc(cols_v[p["col_pivo"]])] = \
                            "background-color: rgba(255,80,0,.55); font-weight: bold"
                return est

            st.dataframe(df.style.apply(destacar, axis=None), hide_index=True, width='stretch')

            if p["status"] == "continua":
                st.markdown(
                    f"🧭 **Entra:** {p['entra']} (maior Cj − Zj = {fmt(p['cj_zj'][p['col_pivo']])}) · "
                    f"🛡️ **Sai:** {p['sai']} (menor razão = {fmt(p['razoes'][p['lin_pivo']])}) · "
                    f"🎯 **Pivô:** {fmt(p['pivo'])}")
                st.code("\n".join(p["operacoes"]), language=None)
            elif p["status"] == "otimo":
                st.markdown("✅ Todos os Cj − Zj ≤ 0: solução ótima.")
            else:
                st.error("Nenhuma razão válida: problema ilimitado.")

    if sol["status"] == "otimo":
        st.divider()
        st.subheader("📊 Business Decision Dashboard")
        v = sol["valores"]
        m1, m2, m3 = st.columns(3)
        m1.metric("X1 · Microsserviços", fmt(v["X1"]))
        m2.metric("X2 · Nós analíticos", fmt(v["X2"]))
        m3.metric("Z máximo (RPS)", fmt(sol["z"]))

        rows = []
        for i, rec in enumerate(restr["Recurso"]):
            folga = v[f"X{3+i}"]
            usado = Fraction(b[i]) - folga
            rows.append({"Recurso": rec, "Disponível": b[i], "Usado": float(usado), "Folga": float(folga),
                         "Situação": "🔴 Gargalo (folga = 0)" if folga == 0 else "🟢 Capacidade ociosa"})
        gf = pd.DataFrame(rows)
        st.dataframe(gf, hide_index=True, width='stretch')
        barras = go.Figure([go.Bar(name="Usado", x=gf["Recurso"], y=gf["Usado"] / gf["Disponível"] * 100),
                            go.Bar(name="Folga", x=gf["Recurso"], y=gf["Folga"] / gf["Disponível"] * 100)])
        barras.update_layout(barmode="stack", yaxis_title="% do recurso", height=340, margin=dict(t=20))
        st.plotly_chart(barras, width='stretch')

        with st.expander("Validação cruzada com SciPy (linprog)"):
            try:
                from scipy.optimize import linprog
                r = linprog([-c_x1, -c_x2], A_ub=A, b_ub=b, bounds=[(0, None)] * 2, method="highs")
                ok = abs(-r.fun - float(sol["z"])) < 1e-6
                st.write(f"SciPy: X1 = {r.x[0]:.4f}, X2 = {r.x[1]:.4f}, Z = {-r.fun:.4f} → "
                         + ("✅ confere com o Simplex próprio" if ok else "⚠️ diverge"))
            except ImportError:
                st.info("Instale scipy para validar.")

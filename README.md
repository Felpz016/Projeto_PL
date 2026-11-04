# Cloud Search · Sistema de Apoio à Decisão (AHP + Simplex)

[![Abrir no Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://cloudsearch-po.streamlit.app)

**Dashboard online:** https://cloudsearch-po.streamlit.app (abre no navegador, sem instalar nada)

Projeto final da disciplina **Introdução à Pesquisa Operacional para Tomada de Decisão**
(Fatec Araraquara · ADS · Prof. Dr. Carlos Ricardo Bifi).

O sistema resolve duas decisões encadeadas de uma startup fictícia, a **Cloud Search**:

1. **Etapa 1 — AHP de Saaty + Normalização Min-Max:** escolhe o provedor de nuvem entre AWS, Azure, Google Cloud,
   Oracle Cloud, DigitalOcean e Linode/Akamai, com 5 critérios (custo, latência, SLA, IOPS e suporte) e dados reais
   coletados em 04/11/2026.
2. **Etapa 2 — Programação Linear pelo Simplex Tabular:** com os preços do provedor vencedor, decide quantas instâncias
   de microsserviços (X1) e de nós analíticos (X2) contratar para maximizar as requisições por segundo.

| Resultado | Valor |
|---|---|
| Provedor recomendado | Oracle Cloud (pontuação 0,8354; RC = 0,0124) |
| Solução ótima | X1 = 24, X2 = 4 |
| Capacidade máxima | Z = 16.800 req/s |
| Gargalos | vCPU e RAM (folga 0); sobra de R$ 884/mês no orçamento |

## Como rodar no seu computador

```bash
python -m venv .venv
.venv\Scripts\activate          # Windows  (Linux/macOS: source .venv/bin/activate)
pip install -r requirements.txt
streamlit run app.py            # abre o dashboard no navegador
python gerar_relatorio.py       # gera relatorio/Relatorio_Projeto_Final_PO.docx
```

## Estrutura

| Arquivo | Função |
|---|---|
| `app.py` | Dashboard Streamlit com as abas Dados, AHP e Simplex |
| `motores/ahp.py` | Normalização Min-Max, pesos de Saaty, λmax, IC, RC e ranking |
| `motores/simplex.py` | Simplex Tabular em frações exatas, com cada quadro e as operações de Gauss-Jordan |
| `motores/precos.py` | Preço ao vivo pela Azure Retail Prices API (pública, sem chave) |
| `dados/pesquisa_mercado.csv` | Pesquisa de mercado com fontes oficiais e status de verificação |
| `gerar_relatorio.py` | Gera o relatório Word a partir dos mesmos motores do dashboard |
| `relatorio/` | Relatório final (.docx) |

## Uso de IA

O projeto foi desenvolvido com apoio do Claude (Anthropic). Todas as interações, os scripts gerados e as validações
estão documentados no Apêndice A do relatório, conforme a Seção 6 da proposta.

**Autor:** Felipe Benaglia Nardini — RA 2880482421038 · Dashboard: https://cloudsearch-po.streamlit.app · Repositório: https://github.com/Felpz016/Projeto_PL

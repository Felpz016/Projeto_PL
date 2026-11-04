# Dashboard de Apoio à Decisão (AHP + Simplex)

    pip install -r requirements.txt
    streamlit run app.py

- `motores/ahp.py`: normalização Min-Max (custo/benefício), pesos de Saaty, λmax, IC, RC e ranking.
- `motores/simplex.py`: Simplex tabular em frações exatas, guardando cada quadro, Cj − Zj, razão mínima e operações de Gauss-Jordan.
- `motores/precos.py`: preço ao vivo pela Azure Retail Prices API (pública, sem chave).
- `app.py`: interface com as abas Dados, AHP e Simplex.

Os dados de mercado em `app.py` (DADOS_EXEMPLO) são ilustrativos. Troque pela pesquisa real da equipe.

## Relatório Word

    python gerar_relatorio.py   # gera relatorio/Relatorio_Projeto_Final_PO.docx

Trechos em amarelo ([PREENCHER] / [CONFERIR]) são da equipe.

## Pendências (dados)
Veja a coluna Status em dados/pesquisa_mercado.csv. Latência é estimada, o preço e o SLA da Oracle precisam de conferência,
e as notas de suporte e alguns IOPS são preliminares.

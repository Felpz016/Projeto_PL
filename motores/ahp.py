"""Motor da Etapa 1: normalização Min-Max e AHP de Saaty."""
import numpy as np
import pandas as pd

# Índice Randômico (IR) de Saaty para n = 1..10
IR_SAATY = {1: 0.0, 2: 0.0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32, 8: 1.41, 9: 1.45, 10: 1.49}


def normalizar_min_max(df: pd.DataFrame, tipos: dict) -> pd.DataFrame:
    """Normaliza cada coluna em [0, 1].

    tipos: {"criterio": "custo" | "beneficio"}
    Benefício: N = (X - min) / (max - min)
    Custo:     N = (max - X) / (max - min)
    """
    norm = pd.DataFrame(index=df.index)
    for col, tipo in tipos.items():
        x = df[col].astype(float)
        xmin, xmax = x.min(), x.max()
        amplitude = xmax - xmin
        if amplitude == 0:  # todas as alternativas empatadas neste critério
            norm[col] = 1.0
        elif tipo == "custo":
            norm[col] = (xmax - x) / amplitude
        else:
            norm[col] = (x - xmin) / amplitude
    return norm


def matriz_de_triangulo(criterios: list, julgamentos: dict) -> np.ndarray:
    """Monta a matriz recíproca n x n a partir dos julgamentos acima da diagonal.

    julgamentos[(i, j)] = a_ij (i < j); a_ji = 1 / a_ij e a_ii = 1.
    """
    n = len(criterios)
    m = np.ones((n, n))
    for (i, j), v in julgamentos.items():
        m[i, j] = v
        m[j, i] = 1.0 / v
    return m


def calcular_ahp(matriz: np.ndarray) -> dict:
    """Calcula pesos (média das linhas da matriz normalizada) e consistência."""
    n = matriz.shape[0]
    soma_colunas = matriz.sum(axis=0)
    normalizada = matriz / soma_colunas
    pesos = normalizada.mean(axis=1)

    # λmax = média de (A·w)_i / w_i
    aw = matriz @ pesos
    lambda_max = float(np.mean(aw / pesos))
    ic = (lambda_max - n) / (n - 1) if n > 1 else 0.0
    ir = IR_SAATY.get(n, 1.49)
    rc = ic / ir if ir > 0 else 0.0
    return {
        "soma_colunas": soma_colunas,
        "normalizada": normalizada,
        "pesos": pesos,
        "aw": aw,
        "lambda_max": lambda_max,
        "ic": ic,
        "ir": ir,
        "rc": rc,
        "consistente": rc < 0.10,
    }


def ranking(norm: pd.DataFrame, pesos: np.ndarray) -> pd.DataFrame:
    """Pontuação_i = Σ W_j * N_ij, ordenada do 1º ao último."""
    pont = norm.values @ pesos
    out = pd.DataFrame({"Pontuação": pont}, index=norm.index)
    out = out.sort_values("Pontuação", ascending=False)
    out.insert(0, "Posição", range(1, len(out) + 1))
    return out

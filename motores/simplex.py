"""Motor da Etapa 2: Simplex Tabular com Gauss-Jordan em frações exatas.

Resolve  Max Z = c·x  s.a.  A x <= b,  x >= 0,  b >= 0
e guarda cada quadro, a variável que entra, o teste da razão mínima,
a variável que sai e as operações de linha aplicadas.
"""
from fractions import Fraction as F


def _fmt(v: F) -> str:
    return str(v.numerator) if v.denominator == 1 else f"{v.numerator}/{v.denominator}"


def resolver(c, A, b, nomes_x=None, nomes_folga=None, max_iter=50):
    m, n = len(A), len(c)
    nomes_x = nomes_x or [f"X{j+1}" for j in range(n)]
    nomes_folga = nomes_folga or [f"X{n+i+1}" for i in range(m)]
    variaveis = list(nomes_x) + list(nomes_folga)

    # Forma padrão: [A | I] x = b
    T = [[F(A[i][j]) for j in range(n)] + [F(1 if k == i else 0) for k in range(m)] + [F(b[i])]
         for i in range(m)]
    custos = [F(v) for v in c] + [F(0)] * m
    base = list(range(n, n + m))  # folgas formam a Matriz Identidade inicial

    passos = []
    for it in range(max_iter + 1):
        cb = [custos[j] for j in base]
        zj = [sum(cb[i] * T[i][j] for i in range(m)) for j in range(n + m)]
        cj_zj = [custos[j] - zj[j] for j in range(n + m)]
        z = sum(cb[i] * T[i][-1] for i in range(m))

        passo = {
            "iteracao": it,
            "variaveis": variaveis,
            "base": [variaveis[j] for j in base],
            "cb": cb,
            "tabela": [linha[:] for linha in T],
            "cj": custos,
            "zj": zj,
            "cj_zj": cj_zj,
            "z": z,
        }

        maior = max(cj_zj)
        if maior <= 0:  # critério de parada
            passo["status"] = "otimo"
            passos.append(passo)
            break

        col = cj_zj.index(maior)
        razoes = [(T[i][-1] / T[i][col]) if T[i][col] > 0 else None for i in range(m)]
        validas = [(r, i) for i, r in enumerate(razoes) if r is not None]
        passo["entra"] = variaveis[col]
        passo["col_pivo"] = col
        passo["razoes"] = razoes
        if not validas:
            passo["status"] = "ilimitado"
            passos.append(passo)
            break

        _, lin = min(validas)
        piv = T[lin][col]
        passo["sai"] = variaveis[base[lin]]
        passo["lin_pivo"] = lin
        passo["pivo"] = piv

        ops = [f"Nova L{lin+1} ({variaveis[col]}) = L{lin+1} antiga ÷ {_fmt(piv)}"]
        T[lin] = [v / piv for v in T[lin]]
        for i in range(m):
            if i != lin and T[i][col] != 0:
                coef = T[i][col]
                ops.append(f"Nova L{i+1} = L{i+1} antiga − ({_fmt(coef)}) × Nova L{lin+1}")
                T[i] = [T[i][j] - coef * T[lin][j] for j in range(n + m + 1)]
        passo["operacoes"] = ops
        passo["status"] = "continua"
        passos.append(passo)
        base[lin] = col

    final = passos[-1]
    valores = {v: F(0) for v in variaveis}
    for i, j in enumerate(base):
        valores[variaveis[j]] = T[i][-1]
    return {"passos": passos, "valores": valores, "z": final["z"], "status": final["status"],
            "nomes_x": nomes_x, "nomes_folga": nomes_folga}


fmt = _fmt

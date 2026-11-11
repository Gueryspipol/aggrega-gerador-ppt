"""Leitor da Matriz EMPRESARIAL (layout de cenários diferente da PME).

Diferenças em relação à PME:
- cada plano ocupa UMA coluna (colunas "VIDAS" vizinhas); na PME são duas (vidas + valor);
- a situação atual traz as faixas etárias com vidas e, abaixo, "Total", "Custo Médio" e "Total Planos";
- as operadoras NÃO trazem valor por faixa etária: só "Subtotal vidas", "Custo Médio" e
  "Subtotal Fatura" (e, quando há desconto/diluição, "Custo Médio c desconto" ou "Diluir fatura" e
  "Subtotal Fatura c desc");
- aba de cenário oculta (ex.: CC quando o cliente não tem coparticipação) é ignorada.

O resultado usa os mesmos modelos da PME (Bloco/Plano/CenarioMatriz), então cálculo e slides são os mesmos.
"""
from __future__ import annotations

import re
from typing import Optional

from .extract import (_PCT, _PLACEHOLDER, _RE_ATUAL, _RE_VIDAS, _Grade, _ancoras_operadora,
                      _num, _RE_OPERADORA, _tem_blocos_operadora, norm)
from .models import Bloco, CenarioMatriz, Plano


# ------------------------------- tipo de Matriz ------------------------------- #
def detectar_tipo(wb) -> Optional[str]:
    """'empresarial' se algum cabeçalho tem colunas VIDAS lado a lado (um plano por coluna);
    'pme' se os cabeçalhos de vidas são separados; None se não achar blocos de operadora."""
    achou = False
    for nome in wb.sheetnames:
        ws = wb[nome]
        if not _tem_blocos_operadora(ws):
            continue
        achou = True
        g = _Grade(ws, max_linhas=120)
        for r, c, v in g.celulas():
            if _RE_VIDAS.match(norm(v)) and _RE_VIDAS.match(norm(g(r, c + 1))):
                return "empresarial"
    return "pme" if achou else None


def conferir_tipo(wb, escolhido: str) -> None:
    detectado = detectar_tipo(wb)
    if detectado and detectado != escolhido:
        raise ValueError(
            f"Este arquivo parece ser uma Matriz {detectado.upper()}, mas foi escolhida a opção "
            f"{escolhido.upper()}. Troque o tipo de Matriz e envie novamente.")


# ------------------------------- leitura de um cenário ------------------------------- #
def _linha(g: _Grade, c0: int, r1: int, r2: int, *rotulos: str, prefixo: bool = False) -> Optional[int]:
    """Primeira linha em [r1, r2) cujo rótulo (coluna c0) é igual (ou começa com) um dos rótulos."""
    for rr in range(r1, r2):
        lab = norm(g(rr, c0)).rstrip("* ").strip()
        if any(lab == x or (prefixo and lab.startswith(x)) for x in rotulos):
            return rr
    return None


def _ler_bloco_emp(g: _Grade, r: int, c0: int, limite: int) -> Bloco:
    c_fim = next((c for c in range(c0 + 1, g.ncols + 1) if _RE_OPERADORA.match(norm(g(r, c)))),
                 g.ncols + 1) - 1
    nome = next((str(g(r, c)).strip() for c in range(c0 + 1, c_fim + 1)
                 if g(r, c) is not None and str(g(r, c)).strip()), "")

    linha_cab = next((rr for rr in range(r + 1, min(r + 7, limite))
                      if any(_RE_VIDAS.match(norm(g(rr, c))) for c in range(c0 + 1, c_fim + 1))), None)
    if linha_cab is None:
        raise ValueError(f"Bloco '{nome}' (linha {r}): cabeçalho com 'VIDAS' não encontrado")
    colunas = [c for c in range(c0 + 1, c_fim + 1) if _RE_VIDAS.match(norm(g(linha_cab, c)))]
    linha_planos = linha_cab - 1 if linha_cab - 1 > r else None

    l_total_planos = _linha(g, c0, linha_cab + 1, limite, "total planos")
    if l_total_planos is not None:                       # situação atual / reajuste
        l_vidas = _linha(g, c0, linha_cab + 1, limite, "total")
        l_valor, l_desc = l_total_planos, None
        faixas = [str(g(rr, c0) or "").strip() for rr in range(linha_cab + 1, l_vidas or linha_cab + 1)]
    else:                                                # proposta de operadora
        l_vidas = _linha(g, c0, linha_cab + 1, limite, "subtotal vidas")
        l_valor = _linha(g, c0, linha_cab + 1, limite, "subtotal fatura")
        l_desc = _linha(g, c0, linha_cab + 1, limite, "subtotal fatura c desc", prefixo=True)
        faixas = []
    if l_vidas is None or l_valor is None:
        raise ValueError(f"Bloco '{nome}' (linha {r}): não achei as linhas de vidas e de valor "
                         f"(layout Empresarial). Confira se escolheu o tipo certo de Matriz.")

    planos, tem_desc = [], False
    for c in colunas:
        rotulo = g(linha_planos, c) if linha_planos else None
        nome_plano = None if (rotulo in (None, "") or _PLACEHOLDER.match(norm(rotulo))
                              or norm(rotulo) == "plano") else str(rotulo).strip()
        vidas = int(_num(g(l_vidas, c)) or 0)
        bruto = _num(g(l_valor, c)) or 0.0
        desc = _num(g(l_desc, c)) if l_desc else None
        if desc and desc > 0:
            tem_desc = True
            planos.append(Plano(nome=nome_plano, vidas=vidas, total=desc, total_bruto=bruto))
        else:
            planos.append(Plano(nome=nome_plano, vidas=vidas, total=bruto))
    return Bloco(indice=0, nome=nome, linha=r, faixas=faixas, planos=planos, com_desconto=tem_desc)


def ler_cenario_empresarial(ws, tipo: str) -> CenarioMatriz:
    g = _Grade(ws)

    def _valida(r: int, c: int) -> bool:
        return any(_RE_VIDAS.match(norm(g(rr, cc)))
                   for rr in range(r + 1, r + 7) for cc in range(c + 1, c + 12))
    ancoras = sorted(a for a in _ancoras_operadora(g) if _valida(*a))
    if not ancoras:
        raise ValueError(f"Aba {ws.title!r}: nenhum bloco 'OPERADORA' com cabeçalho de vidas")

    linhas = sorted({r for r, _ in ancoras})
    lidos = [_ler_bloco_emp(g, r, c, next((x for x in linhas if x > r), g.n + 1)) for r, c in ancoras]

    pos_atual = next((i for i, b in enumerate(lidos) if _RE_ATUAL.search(norm(b.nome))), 0)
    atual = lidos[pos_atual]
    pos_reaj = next((i for i, b in enumerate(lidos)
                     if i != pos_atual and b.linha == atual.linha and "reajust" in norm(b.nome)), None)
    reajuste_bloco = lidos[pos_reaj] if pos_reaj is not None else None
    # blocos sem nome (áreas reservadas do modelo) não são operadoras
    outros = [b for i, b in enumerate(lidos) if i not in (pos_atual, pos_reaj) and b.nome.strip()]
    blocos = [atual] + outros
    for i, b in enumerate(blocos):
        b.indice = i

    reajuste = None
    for _, _, v in g.celulas(r2=60):
        if "reajuste" in norm(v):
            m = _PCT.search(str(v))
            if m:
                reajuste = float(m.group(1).replace(",", ".")) / 100
                break
    fat_atual = atual.total_mensal or None
    fat_reaj = reajuste_bloco.total_mensal if reajuste_bloco else None
    return CenarioMatriz(tipo=tipo, blocos=blocos, reajuste_pct=reajuste, reajuste_bloco=reajuste_bloco,
                         fatura_atual=fat_atual, fatura_reajustada=fat_reaj)

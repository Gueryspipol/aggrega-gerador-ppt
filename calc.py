"""Camada de transformação: replica em Python as regras do Excel de Gráficos.

Regras (conferidas contra as abas 'Análise Financeira' e 'Impacto Financeiro'):
- reajustado = custo médio atual * (1 + reajuste)
- total do plano = custo médio * nº de vidas do plano ATUAL
- impacto mensal = total da operadora - fatura reajustada
- impacto %     = total da operadora / fatura reajustada - 1
- ordem das operadoras: do menor para o maior custo no CC; o SC segue a MESMA ordem, mesmo que no SC
  a classificação mude.
"""
from __future__ import annotations

from typing import Optional, Sequence, Union

from .models import (CenarioCalc, CenarioMatriz, Estudo, Matriz, OperadoraCalc,
                     PlanoAtual, PlanoOperadora)

MAX_OPERADORAS = 4
Selecao = Sequence[Union[int, str]]


def _nome_operadora(matriz: Matriz, indice: int) -> str:
    for tipo in ("CC", "SC"):
        cen = matriz.cenarios.get(tipo)
        if cen and indice < len(cen.blocos) and cen.blocos[indice].nome:
            return cen.blocos[indice].nome
    return f"Operadora {indice}"


def listar_operadoras(matriz: Matriz) -> list[dict]:
    """Operadoras disponíveis (índice 1..n, sem a situação atual) com totais por cenário."""
    n = max(len(c.blocos) for c in matriz.cenarios.values())
    saida = []
    for i in range(1, n):
        item = {"indice": i, "nome": _nome_operadora(matriz, i)}
        for tipo, cen in matriz.cenarios.items():
            item[tipo] = cen.blocos[i].total_mensal if i < len(cen.blocos) else 0.0
        saida.append(item)
    return saida


def resolver_selecao(matriz: Matriz, selecao: Optional[Selecao]) -> list[int]:
    ops = listar_operadoras(matriz)
    por_indice = {o["indice"]: o for o in ops}
    if not selecao:
        escolhidas = [o["indice"] for o in ops
                      if any(o.get(t, 0) > 0 for t in matriz.cenarios)]
        if len(escolhidas) > MAX_OPERADORAS:
            raise ValueError(
                f"A Matriz tem {len(escolhidas)} operadoras com valores; selecione até "
                f"{MAX_OPERADORAS}. Disponíveis: "
                + ", ".join(f"{o['indice']}={o['nome']}" for o in ops if o['indice'] in escolhidas))
        return escolhidas

    resultado: list[int] = []
    for item in selecao:
        if isinstance(item, int) or str(item).strip().isdigit():
            idx = int(item)
            if idx not in por_indice:
                raise ValueError(f"Operadora de índice {idx} não existe (1..{max(por_indice)})")
        else:
            achadas = [o["indice"] for o in ops if str(item).strip().lower() in o["nome"].lower()]
            if len(achadas) != 1:
                raise ValueError(f"Nome '{item}' casa com {len(achadas)} operadoras; use o índice")
            idx = achadas[0]
        if idx not in resultado:
            resultado.append(idx)
    if not 1 <= len(resultado) <= MAX_OPERADORAS:
        raise ValueError(f"Selecione de 1 a {MAX_OPERADORAS} operadoras (recebi {len(resultado)})")
    return resultado


def _separar_reajuste(ativos: list) -> tuple[list, Optional[list]]:
    """Na situação atual, a Matriz às vezes traz o plano e logo depois o MESMO plano com o valor
    reajustado (ex.: S1500, Nacional III, S1500, Nacional III). Se a 2ª metade repete nome e vidas
    da 1ª com custo maior, a 1ª é o ATUAL e a 2ª é o REAJUSTE. Devolve (atual, reajuste|None)."""
    n = len(ativos)
    if n >= 2 and n % 2 == 0:
        h = n // 2
        a, b = ativos[:h], ativos[h:]
        if all(x.nome and x.nome == y.nome and x.vidas == y.vidas and y.custo_medio > x.custo_medio
               for x, y in zip(a, b)):
            return a, b
    return ativos, None


def _calc_cenario(cen: CenarioMatriz, indices: list[int], nome_op, avisos: list[str]) -> CenarioCalc:
    base = cen.blocos[0]
    ativos = [p for p in base.planos if p.vidas > 0]
    if not ativos:
        raise ValueError(f"Cenário {cen.tipo}: nenhum plano atual com vidas")
    atuais_m, reaj_m = ativos, None
    rb = cen.reajuste_bloco
    if rb:       # valores reajustados explícitos (cada plano pode ter % diferente)
        rp = [p for p in rb.planos if p.vidas > 0]
        if len(rp) == len(ativos) and all(x.vidas == y.vidas for x, y in zip(ativos, rp)):
            reaj_m = rp
    explicito = reaj_m is not None
    if reaj_m is None:
        atuais_m, reaj_m = _separar_reajuste(ativos)
    if reaj_m and not explicito:
        avisos.append(f"{cen.tipo}: a situação atual traz {len(ativos)} planos; os {len(reaj_m)} últimos "
                      f"repetem os {len(atuais_m)} primeiros com valor maior e foram tratados como REAJUSTE")

    fatura_atual = sum(p.total for p in atuais_m)
    reaj = cen.reajuste_pct
    if reaj_m:
        fatura_reaj = sum(p.total for p in reaj_m)
        if reaj is None:
            reaj = fatura_reaj / fatura_atual - 1 if fatura_atual else 0.0
    else:
        if reaj is None:
            reaj = (cen.fatura_reajustada / cen.fatura_atual - 1) if cen.fatura_atual and cen.fatura_reajustada else 0.0
            avisos.append(f"{cen.tipo}: percentual de reajuste não encontrado no texto; calculado = {reaj:.4%}")
        fatura_reaj = None
    if cen.fatura_atual and abs(cen.fatura_atual - fatura_atual) > 0.01:
        avisos.append(f"{cen.tipo}: 'Fatura Atual' da Matriz ({cen.fatura_atual:,.2f}) difere da soma dos "
                      f"planos atuais ({fatura_atual:,.2f})")

    atuais = []
    for i, p in enumerate(atuais_m):
        cm = p.custo_medio
        cm_r = reaj_m[i].custo_medio if reaj_m else cm * (1 + reaj)
        atuais.append(PlanoAtual(nome=p.nome or f"Plano {i + 1}", vidas=p.vidas, custo_medio=cm,
                                 total=p.total, custo_medio_reaj=cm_r, total_reaj=cm_r * p.vidas))
    fatura_reaj = sum(a.total_reaj for a in atuais)

    ops = []
    for idx in indices:
        if idx >= len(cen.blocos):
            avisos.append(f"{cen.tipo}: operadora {idx} ausente na aba")
            continue
        planos = [PlanoOperadora(nome=p.nome or f"Plano {k + 1}", vidas=p.vidas, custo_medio=p.custo_medio,
                                 total=p.total, por_faixa=p.por_faixa)
                  for k, p in enumerate(cen.blocos[idx].planos) if p.vidas > 0]
        if not planos:
            avisos.append(f"{cen.tipo}: {nome_op(idx)} sem nenhum plano com valores")
        total = sum(p.total for p in planos)
        bloco = cen.blocos[idx]
        if bloco.com_desconto:
            bruto = sum(p.total_bruto if p.total_bruto is not None else p.total for p in bloco.planos if p.vidas > 0)
            avisos.append(f"{cen.tipo}: {nome_op(idx)} usa o valor COM desconto/diluição da fatura "
                          f"({total:,.2f}); sem desconto seria {bruto:,.2f}")
        ops.append(OperadoraCalc(
            nome=nome_op(idx), bloco_indice=idx, total_mensal=total,
            reducao_mensal=total - fatura_reaj,
            reducao_pct=(total / fatura_reaj - 1) if fatura_reaj else 0.0,
            planos=planos))

    return CenarioCalc(tipo=cen.tipo, reajuste_pct=reaj, fatura_atual=fatura_atual,
                       fatura_reajustada=fatura_reaj, total_vidas=sum(a.vidas for a in atuais),
                       faixas=base.faixas, planos_atuais=atuais, operadoras=ops)


def calcular(matriz: Matriz, selecao: Optional[Selecao] = None, *,
             cliente: Optional[str] = None,
             mes_atual: Optional[tuple[str, int]] = None,
             mes_reajuste: Optional[tuple[str, int]] = None,
             operadora_atual: Optional[str] = None) -> Estudo:
    avisos: list[str] = list(matriz.avisos)
    indices = resolver_selecao(matriz, selecao)

    # Ordem das colunas: do MENOR para o MAIOR custo no CC (senão SC). O SC usa exatamente a mesma
    # ordem, mesmo que lá o ranking mude.
    ref = matriz.cenarios.get("CC") or matriz.cenarios["SC"]
    indices.sort(key=lambda i: (ref.blocos[i].total_mensal if i < len(ref.blocos) else float("inf"), i))

    nome_op = lambda i: _nome_operadora(matriz, i)  # noqa: E731
    cenarios = {}
    for tipo, cen in matriz.cenarios.items():
        if len(cen.blocos) != len(ref.blocos):
            avisos.append(f"{tipo}: número de blocos ({len(cen.blocos)}) difere do cenário de referência ({len(ref.blocos)})")
        cenarios[tipo] = _calc_cenario(cen, indices, nome_op, avisos)

    if "SC" not in cenarios:
        avisos.append("Matriz sem aba SC: apenas o cenário CC será preenchido")
    return Estudo(cliente=cliente or matriz.cliente or "", mes_atual=mes_atual,
                  mes_reajuste=mes_reajuste, demografia=matriz.demografia,
                  cenarios=cenarios, avisos=avisos,
                  operadora_atual=(operadora_atual or "").strip()
                  or " ".join(x for i, x in enumerate((ref.blocos[0].nome, matriz.congenere))
                              if x and not (i == 1 and x.lower() in ref.blocos[0].nome.lower())))

"""Camada de carga: preenche os slides do template (1, 13, 14-17 CC e 18-21 SC).

Os elementos são localizados pelo que são (texto, posição, tipo) e não por nomes ou ids,
assim o mesmo código serve aos slides CC e SC e sobrevive a pequenas mudanças no template.
Nada de layout é alterado: só texto, células de tabela e dados de gráficos nativos.
"""
from __future__ import annotations

import re
from typing import Optional

from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.dml.color import RGBColor
from pptx.util import Pt

from . import formatos as f
from .extract import FAIXA_INICIOS, FAIXA_ROTULOS
from .models import CenarioCalc, Demografia, Estudo, OperadoraCalc
from .pptx_writer import (clonar_caixa, definir_celula, definir_linhas, definir_paragrafo,
                          duplicar_slide, limpar, substituir_dados_grafico)

# número do slide (1-based) de cada parte do template
SLIDE_CAPA = 1
SLIDE_DEMOGRAFIA = 13
SLIDES_CENARIO = {
    "CC": {"resumo": 14, "financeira": 15, "faixas": [16, 17]},
    "SC": {"resumo": 18, "financeira": 19, "faixas": [20, 21]},
}

_RE_VALOR = re.compile(r"^-?R\$\s*[\d.,]*$")   # inclui o "R$ " vazio do template
_RE_BADGE = re.compile(r"^[<>]?\s*[+-]?\d+(?:[.,]\d+)?%$")   # "-70%", "< 39,53%"
_RE_MES = re.compile(r"^[A-ZÇ]{3}$")


def _caixas_texto(slide):
    return [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]


# --------------------------------------------------------------------------- #
# Slide 1 – capa
# --------------------------------------------------------------------------- #
def preencher_capa(slide, cliente: str) -> None:
    alvo = max(_caixas_texto(slide), key=lambda s: s.top)   # nome do cliente: caixa mais baixa
    definir_linhas(alvo, [cliente])


# --------------------------------------------------------------------------- #
# Slide 13 – painel demográfico
# --------------------------------------------------------------------------- #
def _pct_inteiros(partes: list[int]) -> list[int]:
    total = sum(partes)
    if not total:
        return [0] * len(partes)
    pct = [int(round(p * 100 / total + 1e-9)) for p in partes]
    pct[-1] += 100 - sum(pct)          # garante soma 100
    return pct


def preencher_demografia(slide, demo: Demografia, cen: CenarioCalc) -> None:
    for sh in slide.shapes:
        if not getattr(sh, "has_chart", False) or not sh.has_chart:
            continue
        chart = sh.chart
        tipo = chart.chart_type
        nomes_series = [s.name.upper() for s in chart.plots[0].series]
        if "DOUGHNUT" in str(tipo):
            substituir_dados_grafico(chart, list(demo.titularidade),
                                     [list(demo.titularidade.values())])
        elif "PIE" in str(tipo):
            vidas = [a.vidas for a in cen.planos_atuais]
            total = sum(vidas) or 1
            substituir_dados_grafico(chart, [a.nome for a in cen.planos_atuais],
                                     [vidas, [v / total for v in vidas]])
        elif "FUNCIONÁRIOS" in nomes_series:
            t = [demo.faixa_etaria[i]["T"] or None for i in FAIXA_INICIOS]
            d = [demo.faixa_etaria[i]["D"] or None for i in FAIXA_INICIOS]
            substituir_dados_grafico(chart, FAIXA_ROTULOS, [t, d])
        else:   # distribuição do plano: valor atual por plano
            cats = [f"{a.nome}\n{a.vidas} VIDAS" for a in cen.planos_atuais]
            substituir_dados_grafico(chart, cats, [[a.total for a in cen.planos_atuais],
                                                   [None] * len(cats)])

    caixas = _caixas_texto(slide)
    masc = next((s for s in caixas if s.text_frame.text.strip().startswith("Masculino")), None)
    fem = next((s for s in caixas if s.text_frame.text.strip().startswith("Feminino")), None)
    m, fe = demo.genero["M"], demo.genero["F"]
    if masc is not None and fem is not None and (m + fe):
        definir_linhas(masc, ["Masculino", f"({m})"])
        definir_linhas(fem, [f"Feminino ({fe})"])
        for sh in caixas:
            t = sh.text_frame.text.strip()
            if re.fullmatch(r"\d+(?:[.,]\d+)?%", t):
                casas = 1 if re.search(r"[.,]\d", t) else 0        # mantém o formato do template
                pm = round(m * 100 / (m + fe), casas)
                valor = pm if abs(sh.left - masc.left) < abs(sh.left - fem.left) else round(100 - pm, casas)
                texto = f"{valor:.1f}".replace(".", ",") if casas else f"{int(valor)}"
                definir_linhas(sh, [texto + "%"])


# --------------------------------------------------------------------------- #
# Slides 14 / 18 – resumo executivo
# --------------------------------------------------------------------------- #
def _colunas_resumo(slide, altura_slide: int):
    """Cartões (colunas) do slide de resumo e as caixas de texto de cada coluna."""
    cartoes = sorted((s for s in slide.shapes
                      if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                      and s.height >= 0.3 * altura_slide and not s.text_frame.text.strip()),
                     key=lambda s: s.left)
    if len(cartoes) < 3:
        raise ValueError("Slide de resumo: cartões de coluna não encontrados no template")
    por_coluna: dict[int, list] = {i: [] for i in range(len(cartoes))}
    for sh in _caixas_texto(slide):
        if not (0.15 * altura_slide < sh.top < 0.92 * altura_slide):
            continue                                        # título e rodapés ficam intactos
        cx = sh.left + sh.width / 2
        for i, c in enumerate(cartoes):
            if c.left <= cx <= c.left + c.width:
                por_coluna[i].append(sh)
                break
    return cartoes, por_coluna


def _reatribuir_etiquetas(por_coluna, n_cols: int, n_usadas: int) -> None:
    """As operadoras entram do menor para o maior custo, então a etiqueta "Menor Custo" vai para a
    1ª coluna e as demais etiquetas do template (ex.: "+ Opções", "Upgrade Produto") seguem na ordem."""
    etiquetas = [min(por_coluna[c + 2], key=lambda s: s.top) if por_coluna[c + 2] else None
                 for c in range(n_cols)]
    if any(e is None for e in etiquetas):
        return
    textos = [e.text_frame.text for e in etiquetas]
    if not any("menor custo" in t.lower() for t in textos):
        return
    novos = ["Menor Custo"] + [t for t in textos if "menor custo" not in t.lower()]
    for j in range(min(n_usadas, n_cols)):
        if novos[j] != textos[j]:
            definir_linhas(etiquetas[j], [linha.strip() for linha in novos[j].split("\n")])


def preencher_resumo(slide, cen: CenarioCalc, estudo: Estudo, altura_slide: int,
                     avisos: list[str]) -> None:
    cartoes, por_coluna = _colunas_resumo(slide, altura_slide)

    def achar(col, teste):
        return [s for s in por_coluna[col] if teste(s.text_frame.text.strip())]

    # colunas 0 (atual) e 1 (reajuste)
    for col, valor, mes in ((0, cen.fatura_atual, estudo.mes_atual),
                            (1, cen.fatura_reajustada, estudo.mes_reajuste)):
        for sh in achar(col, lambda t: bool(_RE_VALOR.match(t))):
            definir_linhas(sh, [f.brl_inteiro(valor)])
        for sh in achar(col, lambda t: t.startswith("Total de Vidas")):
            definir_linhas(sh, ["Total de Vidas", str(cen.total_vidas)])
        for sh in achar(col, lambda t: bool(_RE_MES.match(t.split("\n")[0]))):
            if mes:
                definir_linhas(sh, [mes[0], str(mes[1])])
            else:
                avisos.append(f"Slide de resumo {cen.tipo}: mês/ano da coluna {col + 1} não informado; mantido o do template")
    for sh in achar(1, lambda t: "reajuste" in t.lower()):
        for par in sh.text_frame.paragraphs:                # "Reajuste %" ou "Reajuste 11,83%"
            txt = "".join(r.text for r in par.runs)
            if "reajuste" in txt.lower():
                definir_paragrafo(par, re.sub(r"(?i)(reajuste)\s*(?:\d+(?:[.,]\d+)?)?\s*%?",
                                              lambda m: f"{m.group(1)} {f.pct_2casas(cen.reajuste_pct)}", txt))

    # colunas de operadoras
    n_cols = len(cartoes) - 2
    if len(cen.operadoras) > n_cols:
        avisos.append(f"Slide de resumo {cen.tipo}: template tem {n_cols} colunas de operadora; "
                      f"{len(cen.operadoras) - n_cols} operadora(s) não couberam")
    _reatribuir_etiquetas(por_coluna, n_cols, len(cen.operadoras))
    for j in range(n_cols):
        col = j + 2
        op: Optional[OperadoraCalc] = cen.operadoras[j] if j < len(cen.operadoras) else None
        if op is None:
            for sh in por_coluna[col]:
                limpar(sh)
            continue
        for sh in achar(col, lambda t: bool(_RE_VALOR.match(t))):
            definir_linhas(sh, [f.brl_inteiro(op.total_mensal)])
        for sh in achar(col, lambda t: bool(_RE_BADGE.match(t))):
            definir_linhas(sh, [f.pct_inteiro(op.reducao_pct)])
        for sh in achar(col, lambda t: t.startswith("Redução")):
            definir_linhas(sh, ["Redução Mensal", f.brl_delta(op.reducao_mensal)])


# --------------------------------------------------------------------------- #
# Slides 15 / 19 – análise financeira
# --------------------------------------------------------------------------- #
_ROTULOS_LINHA = ["Produto", "Número de Vidas", "Custo Médio", "Total"]


def _modelo_celula(tabela, k: int, c: int):
    for cc in range(c, -1, -1):
        cel = tabela.cell(k, cc)
        if cel.text_frame.paragraphs[0].runs:
            return cel
    return None


def _colunas_financeira(cen: CenarioCalc, n_ops_cols: int) -> dict[int, list[tuple]]:
    colunas: dict[int, list[tuple]] = {
        1: [(a.nome, a.vidas, a.custo_medio, a.total) for a in cen.planos_atuais],
        2: [(a.nome, a.vidas, a.custo_medio_reaj, a.total_reaj) for a in cen.planos_atuais],
    }
    for j, op in enumerate(cen.operadoras[:n_ops_cols]):
        colunas[3 + j] = [(p.nome, p.vidas, p.custo_medio, p.total) for p in op.planos]
    return colunas


def capacidade_financeira(slide) -> tuple[int, int]:
    """(planos por slide, colunas de operadora) da tabela de análise financeira."""
    tabela = next(s for s in slide.shapes if s.has_table).table
    return (len(tabela.rows) + 1) // 5, len(tabela.columns) - 3


def paginas_financeira(slide, cen: CenarioCalc) -> int:
    """Quantos slides de análise financeira são necessários para listar todos os planos."""
    n_grupos, n_ops_cols = capacidade_financeira(slide)
    maior = max((len(v) for v in _colunas_financeira(cen, n_ops_cols).values()), default=0)
    return max(1, -(-maior // n_grupos))


def preencher_financeira(slide, cen: CenarioCalc, avisos: list[str], pagina: int = 0) -> None:
    """Colunas: Atual | Reajuste | operadoras. Cada coluna lista os seus próprios planos, de cima
    para baixo (blocos de 4 linhas: produto, vidas, custo médio, total). Quando há mais planos do
    que a tabela comporta, os seguintes continuam no slide duplicado (`pagina` 1, 2, ...)."""
    tabela = next(s for s in slide.shapes if s.has_table).table
    n_cols = len(tabela.columns)
    n_grupos, n_ops_cols = capacidade_financeira(slide)
    if pagina == 0 and len(cen.operadoras) > n_ops_cols:
        avisos.append(f"Slide de análise financeira {cen.tipo}: tabela tem {n_ops_cols} colunas de operadora; "
                      f"{len(cen.operadoras) - n_ops_cols} não couberam")

    inicio = pagina * n_grupos
    colunas = {c: lista[inicio:inicio + n_grupos] for c, lista in _colunas_financeira(cen, n_ops_cols).items()}
    usados = max((len(v) for v in colunas.values()), default=0)

    for g in range(n_grupos):
        for k in range(4):
            if 5 * g + k >= len(tabela.rows):
                break
            valores = [""] * n_cols
            if g < usados:
                valores[0] = _ROTULOS_LINHA[k]
            for c, lista in colunas.items():
                if g < len(lista):
                    nome, vidas, custo_medio, total = lista[g]
                    valores[c] = [nome, str(vidas), f.brl_centavos(custo_medio), f.brl_centavos(total)][k]
            for c in range(n_cols):
                definir_celula(tabela.cell(5 * g + k, c), valores[c], _modelo_celula(tabela, k, c))


# --------------------------------------------------------------------------- #
# Logos das operadoras
# --------------------------------------------------------------------------- #
IN = 914400


def _colar_logos(slide, caminhos, cx: float, cy: float, larg: float, alt: float) -> None:
    """Cola um ou mais logos lado a lado, sem distorcer, centralizados em (cx, cy) dentro de larg x alt."""
    from PIL import Image
    caminhos = [c for c in caminhos if c]
    if not caminhos:
        return
    faixa = larg / len(caminhos)
    for i, caminho in enumerate(caminhos):
        with Image.open(caminho) as im:
            w, h = im.size
        esc = min(faixa * 0.92 / w, alt / h)
        pw, ph = int(w * esc), int(h * esc)
        x = int(cx - larg / 2 + faixa * (i + 0.5) - pw / 2)
        slide.shapes.add_picture(str(caminho), x, int(cy - ph / 2), pw, ph)


def logos_resumo(slide, logos_cols: list, altura_slide: int) -> None:
    """Slides 14/18: o logo vai no balão branco que fica acima de cada coluna (cartão)."""
    auto = [x for x in slide.shapes if x.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and not x.text_frame.text.strip()]
    cartoes = sorted((x for x in auto if x.height >= 0.3 * altura_slide), key=lambda x: x.left)
    baloes = [x for x in auto if x.height < 0.3 * altura_slide]
    for caminhos, cartao in zip(logos_cols, cartoes):
        candidatos = [b for b in baloes
                      if cartao.left <= b.left + b.width / 2 <= cartao.left + cartao.width
                      and b.top < cartao.top + 0.2 * IN]
        if not candidatos:
            continue
        balao = min(candidatos, key=lambda b: abs(b.top - cartao.top))
        _colar_logos(slide, caminhos, balao.left + balao.width / 2, balao.top + balao.height * 0.40,
                     balao.width * 0.82, balao.height * 0.50)


def logos_tabela_financeira(slide, logos_cols: list) -> None:
    """Slides 15/19: o logo vai entre as etiquetas (ATUAL, REAJUSTE...) e o topo de cada coluna da tabela."""
    forma = next(x for x in slide.shapes if x.has_table)
    abaixo_de = [x.top + x.height for x in slide.shapes
                 if x.shape_type not in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINE, MSO_SHAPE_TYPE.TABLE)
                 and x.top + x.height < forma.top]
    topo = max(abaixo_de, default=forma.top - 1.0 * IN) + 0.06 * IN
    base = forma.top - 0.04 * IN
    alt = max(0.3 * IN, min(0.65 * IN, (base - topo) * 0.9))
    esquerda = forma.left
    for c, col in enumerate(forma.table.columns):
        if 1 <= c <= len(logos_cols) and logos_cols[c - 1]:
            _colar_logos(slide, logos_cols[c - 1], esquerda + col.width / 2, (topo + base) / 2,
                         col.width * 0.72, alt)
        esquerda += col.width


def logos_faixas(slide, grupos_logos: list[tuple[int, int, list]], altura_slide: int) -> None:
    """Slides 16/17/20/21: o logo da operadora fica dentro do cartão de cada coluna, logo acima do nome
    do plano (grupos_logos = [(coluna_inicial, coluna_final, [logos])], colunas contadas a partir de 0)."""
    largura = slide.part.package.presentation_part.presentation.slide_width
    caixas = [x for x in slide.shapes if x.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
              and 0.5 * altura_slide <= x.height <= 0.8 * altura_slide and x.width < 0.12 * largura]
    caixas.sort(key=lambda x: x.left)
    if not caixas:
        return
    topo = min(x.top for x in caixas)
    for ini, fim, caminhos in grupos_logos:
        if not caminhos or ini >= len(caixas):
            continue
        a, b = caixas[ini], caixas[min(fim, len(caixas) - 1)]
        span = b.left + b.width - a.left
        _colar_logos(slide, caminhos, a.left + span / 2, topo + 0.62 * IN, min(span * 0.80, 1.3 * IN), 0.50 * IN)


# --------------------------------------------------------------------------- #
# Slides 16-17 / 20-21 – custo por faixa etária
# --------------------------------------------------------------------------- #
def preencher_faixas(slide, cen: CenarioCalc, grupo: list[OperadoraCalc],
                     avisos: list[str]) -> None:
    quadro = next(s for s in slide.shapes if s.has_table)
    tabela = quadro.table
    n_valores = len(tabela.columns) - 1
    planos_col = [p for op in grupo for p in op.planos]
    if len(planos_col) > n_valores:
        avisos.append(f"Slide de faixas {cen.tipo}: {len(planos_col)} planos para {n_valores} colunas; "
                      f"{len(planos_col) - n_valores} não couberam")
    alvo = {c: (planos_col[c - 1] if c - 1 < len(planos_col) else None) for c in range(1, n_valores + 1)}

    n_faixas = min(len(cen.faixas), len(tabela.rows))
    for c in range(1, n_valores + 1):
        plano = alvo[c]
        for r in range(len(tabela.rows)):
            valor = ""
            if plano is not None and r < n_faixas and r < len(plano.por_faixa):
                valor = f.brl_centavos_espaco(plano.por_faixa[r])
            definir_celula(tabela.cell(r, c), valor, _modelo_celula(tabela, r, c))

    # rótulos (nome do plano) acima de cada coluna
    x = quadro.left
    centros = {0: x + tabela.columns[0].width / 2}
    for c in range(1, n_valores + 1):
        x_ini = quadro.left + sum(tabela.columns[i].width for i in range(c))
        centros[c] = x_ini + tabela.columns[c].width / 2
    rotulos = [s for s in _caixas_texto(slide)
               if quadro.top - 700000 <= s.top < quadro.top and not s.has_table]
    por_col: dict[int, object] = {}
    for sh in rotulos:
        cx = sh.left + sh.width / 2
        c = min(range(1, n_valores + 1), key=lambda i: abs(centros[i] - cx))
        por_col[c] = sh
    modelo = rotulos[0] if rotulos else None
    caixas_rotulo: dict[int, object] = {}
    for c in range(1, n_valores + 1):
        plano = alvo[c]
        sh = por_col.get(c)
        if sh is None and plano is not None and modelo is not None:
            sh = clonar_caixa(modelo, centros[c])
        if sh is not None:
            definir_linhas(sh, [plano.nome if plano is not None else ""])
            caixas_rotulo[c] = sh
    _padronizar_rotulos(caixas_rotulo, alvo, centros, quadro.top, tabela, cen.tipo, avisos)


LARG_ROTULO = int(1.5 * IN)       # largura útil do nome (um pouco menor que o cartão da coluna)
ALT_ROTULO = int(0.62 * IN)       # cabe duas linhas; o texto fica colado na tabela (base alinhada)
FONTE_ROTULO, FONTE_ROTULO_MIN = 14.0, 10.0
COR_ROTULO = RGBColor(0x00, 0x00, 0x00)   # nomes dos planos em preto (o template usa cinza-escuro)


def _tamanho_rotulo(nomes: list[str]) -> tuple[float, bool]:
    """Maior fonte (<= padrão) em que o nome mais longo cabe em UMA linha; (fonte, precisa_quebrar)."""
    maior = max((len(n) for n in nomes), default=0)
    if maior == 0:
        return FONTE_ROTULO, False
    cabe = (LARG_ROTULO / IN) * 72 / (0.60 * maior)         # 0,60 em ~ largura média do negrito
    tam = max(FONTE_ROTULO_MIN, min(FONTE_ROTULO, cabe))
    return round(tam * 2) / 2, cabe < FONTE_ROTULO_MIN


def _padronizar_rotulos(caixas: dict, alvo: dict, centros: dict, topo_tabela: int, tabela,
                        tipo: str, avisos: list[str]) -> None:
    """Todos os nomes de plano com a mesma largura, mesma base e a mesma fonte, centralizados na coluna."""
    nomes = [alvo[c].nome for c in caixas if alvo.get(c) is not None]
    tam, quebra = _tamanho_rotulo(nomes)
    if quebra:
        longos = [n for n in nomes if len(n) * 0.60 * FONTE_ROTULO_MIN / 72 * IN > LARG_ROTULO]
        avisos.append(f"Slide de faixas {tipo}: nome de plano muito longo ({', '.join(longos[:3])}); "
                      f"ficou em duas linhas, confira o slide")
    for c, sh in caixas.items():
        corpo = sh.text_frame._txBody.find(qn("a:bodyPr"))
        corpo.set("wrap", "square"); corpo.set("anchor", "b")
        corpo.set("lIns", "0"); corpo.set("rIns", "0"); corpo.set("tIns", "0"); corpo.set("bIns", "0")
        for filho in list(corpo):
            corpo.remove(filho)
        corpo.append(corpo.makeelement(qn("a:noAutofit"), {}))
        sh.width, sh.height = LARG_ROTULO, ALT_ROTULO
        sh.left = int(centros[c] - LARG_ROTULO / 2)
        sh.top = int(topo_tabela - 0.04 * IN - ALT_ROTULO)
        for par in sh.text_frame.paragraphs:
            for r in par.runs:
                r.font.size = Pt(tam)
                r.font.color.rgb = COR_ROTULO


# --------------------------------------------------------------------------- #
# Validação do template
# --------------------------------------------------------------------------- #
def _tem(slide, tipo: str) -> int:
    if tipo == "tabela":
        return sum(1 for s in slide.shapes if getattr(s, "has_table", False) and s.has_table)
    if tipo == "grafico":
        return sum(1 for s in slide.shapes if getattr(s, "has_chart", False) and s.has_chart)
    return sum(1 for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip())


def resumo_template(prs, ate: int = 24) -> str:
    """Uma linha por slide: o que o robô enxerga nele (para diagnosticar templates diferentes)."""
    linhas = [f"O template tem {len(prs.slides)} slides."]
    for i, sl in enumerate(prs.slides, 1):
        if i > ate:
            break
        textos = [s.text_frame.text.strip().split("\n")[0] for s in sl.shapes
                  if s.has_text_frame and s.text_frame.text.strip()]
        linhas.append(f"  slide {i}: {_tem(sl, 'tabela')} tabela(s), {_tem(sl, 'grafico')} gráfico(s), "
                      f"{len(textos)} texto(s) | {textos[0][:40] if textos else '(sem texto)'}")
    return "\n".join(linhas)


def validar_template(prs, tipos: list[str]) -> list[str]:
    """Problemas que impedem o preenchimento (lista vazia = template compatível)."""
    slides = list(prs.slides)
    precisa = [SLIDE_CAPA, SLIDE_DEMOGRAFIA] + [n for t in tipos for n in (
        SLIDES_CENARIO[t]["resumo"], SLIDES_CENARIO[t]["financeira"], *SLIDES_CENARIO[t]["faixas"])]
    if len(slides) < max(precisa):
        return [f"o template tem {len(slides)} slides, mas o robô espera ao menos {max(precisa)}"]
    problemas = []
    if _tem(slides[SLIDE_CAPA - 1], "texto") == 0:
        problemas.append(f"slide {SLIDE_CAPA} (capa): não tem caixa de texto para o nome do cliente")
    if _tem(slides[SLIDE_DEMOGRAFIA - 1], "grafico") == 0:
        problemas.append(f"slide {SLIDE_DEMOGRAFIA} (painel demográfico): não tem gráficos")
    for t in tipos:
        m = SLIDES_CENARIO[t]
        altura = prs.slide_height
        cartoes = [s for s in slides[m["resumo"] - 1].shapes
                   if s.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE and s.height >= 0.3 * altura
                   and not s.text_frame.text.strip()]
        if len(cartoes) < 3:
            problemas.append(f"slide {m['resumo']} (resumo {t}): esperava as colunas (cartões) de valores")
        for n, nome in [(m["financeira"], "análise financeira"), *[(x, "custo por faixa etária") for x in m["faixas"]]]:
            if _tem(slides[n - 1], "tabela") == 0:
                problemas.append(f"slide {n} ({nome} {t}): esperava uma tabela e não achei")
    return problemas


# --------------------------------------------------------------------------- #
def preencher_estudo(prs, estudo: Estudo, avisos: list[str], catalogo=None, com_faixas: bool = True) -> list:
    """Preenche todos os slides e devolve a lista de slides a remover."""
    slides = list(prs.slides)
    problemas = validar_template(prs, list(estudo.cenarios))
    if problemas:
        raise ValueError("O template não tem a estrutura esperada:\n- " + "\n- ".join(problemas)
                         + "\n\n" + resumo_template(prs))
    s = lambda n: slides[n - 1]   # noqa: E731
    remover = []

    sem_logo: list[str] = []
    atuais_logos = catalogo.achar_todos(estudo.operadora_atual) if catalogo else []
    if catalogo is not None:
        if not catalogo:
            avisos.append("Pasta de logos não encontrada ou vazia: os slides ficam sem logos das operadoras")
        elif not atuais_logos:
            avisos.append("Logo da operadora atual não identificado"
                          + (f" ({estudo.operadora_atual})" if estudo.operadora_atual else "")
                          + ": informe o nome da operadora atual ou cadastre o logo")

    def _logos_cols(cen):
        if not catalogo:
            return []
        cols = [atuais_logos, atuais_logos]
        for op in cen.operadoras:
            arq = catalogo.achar(op.nome)
            if arq is None and op.nome not in sem_logo:
                sem_logo.append(op.nome)
            cols.append([arq] if arq else [])
        return cols

    preencher_capa(s(SLIDE_CAPA), estudo.cliente)
    primeiro = next(iter(estudo.cenarios.values()))
    preencher_demografia(s(SLIDE_DEMOGRAFIA), estudo.demografia, primeiro)

    for tipo, mapa in SLIDES_CENARIO.items():
        cen = estudo.cenarios.get(tipo)
        todos = [mapa["resumo"], mapa["financeira"], *mapa["faixas"]]
        if cen is None:
            remover += [s(n) for n in todos]
            continue
        logos_cols = _logos_cols(cen)
        preencher_resumo(s(mapa["resumo"]), cen, estudo, prs.slide_height, avisos)
        if logos_cols:
            logos_resumo(s(mapa["resumo"]), logos_cols, prs.slide_height)
        original = s(mapa["financeira"])
        n_pag = paginas_financeira(original, cen)
        paginas = [original]
        for _ in range(n_pag - 1):             # mais planos que linhas: duplica o slide e continua
            paginas.append(duplicar_slide(prs, original, depois=paginas[-1]))
        if n_pag > 1:
            avisos.append(f"Análise financeira {tipo}: planos demais para um slide; "
                          f"o slide foi duplicado e a lista continua em {n_pag} slides")
        for i, pg in enumerate(paginas):
            preencher_financeira(pg, cen, avisos, pagina=i)
            if logos_cols:
                logos_tabela_financeira(pg, logos_cols)
        if not com_faixas:                     # Matriz Empresarial: sem custo por faixa etária
            remover += [s(n) for n in mapa["faixas"]]
            continue
        tabela_faixas = next(x for x in s(mapa["faixas"][0]).shapes if x.has_table).table
        capacidade = len(tabela_faixas.columns) - 1
        grupos: list[list[OperadoraCalc]] = [[]]
        usado = 0
        for op in cen.operadoras:
            if grupos[-1] and usado + len(op.planos) > capacidade:
                grupos.append([])
                usado = 0
            grupos[-1].append(op)
            usado += len(op.planos)
        grupos = [g for g in grupos if g]
        # slides de faixas necessários: usa os do template e duplica o último quando faltar
        modelo = s(mapa["faixas"][-1])
        paginas = [s(n) for n in mapa["faixas"][:len(grupos)]]
        while len(paginas) < len(grupos):
            paginas.append(duplicar_slide(prs, modelo, depois=paginas[-1]))
        if len(grupos) > len(mapa["faixas"]):
            avisos.append(f"Custo por faixa etária {tipo}: planos demais para {len(mapa['faixas'])} slides; "
                          f"foram criados mais {len(grupos) - len(mapa['faixas'])} slide(s)")
        for pg, grupo in zip(paginas, grupos):
            preencher_faixas(pg, cen, grupo, avisos)
            if catalogo:
                colunas, ini = [], 0
                for op in grupo:          # um logo sobre CADA coluna de plano da operadora
                    logo = [a] if (a := catalogo.achar(op.nome)) else []
                    for k in range(len(op.planos)):
                        colunas.append((ini + k, ini + k, logo))
                    ini += len(op.planos)
                logos_faixas(pg, colunas, prs.slide_height)
        remover += [s(n) for n in mapa["faixas"][len(grupos):]]
    if sem_logo:
        avisos.append("Sem logo cadastrado para: " + ", ".join(sem_logo)
                      + " (logos reconhecidos: " + ", ".join(sorted(a.stem for a in (x[1] for x in catalogo.itens))) + ")")
    return remover

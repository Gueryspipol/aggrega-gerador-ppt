"""Testes de regressão contra os arquivos de exemplo (Shock Imports).

Os exemplos têm dados de cliente e NÃO vão para o Git. Aponte a pasta com eles:

    AP360_SAMPLES=/caminho/da/pasta pytest

Arquivos esperados na pasta:
    Shock_Imports_Matriz_PME_2026_06_v1.xlsx   (Matriz)
    Apr_Shock_Imports_v1_062026.pptx           (template/apresentação de referência)
    Apr_Shock_Imports_v1_062026.xlsx           (Excel de Gráficos de referência)
"""
import os
from pathlib import Path

import openpyxl
import pytest
from pptx import Presentation

from apresentador360 import formatos as f
from apresentador360.calc import calcular, resolver_selecao
from apresentador360.extract import achar_aba, carregar_matriz, norm
from apresentador360.pipeline import gerar, mes_do_arquivo, parse_mes

PASTA = Path(os.environ.get("AP360_SAMPLES", "samples"))
MATRIZ = PASTA / "Shock_Imports_Matriz_PME_2026_06_v1.xlsx"
PPTX = PASTA / "Apr_Shock_Imports_v1_062026.pptx"
XLSX = PASTA / "Apr_Shock_Imports_v1_062026.xlsx"
precisa_amostras = pytest.mark.skipif(not (MATRIZ.exists() and PPTX.exists()),
                                      reason="arquivos de exemplo não encontrados (AP360_SAMPLES)")


# ---------- unidades (não precisam de arquivos) ---------------------------- #
def test_formatos():
    assert f.brl_inteiro(18219.6) == "R$18.220"
    assert f.brl_inteiro(6136.08) == "R$6.136"
    assert f.brl_delta(-14238.89868) == "-R$14.239"
    assert f.brl_delta(2155.37) == "+R$2.155"
    assert f.brl_centavos(1152.49) == "R$1.152,49"
    assert f.brl_centavos(12869.12) == "R$12.869,12"
    assert f.brl_centavos_espaco(1029.95) == "R$ 1.029,95"
    assert f.pct_inteiro(-0.6988423842610862) == "-70%"
    assert f.pct_2casas(0.1183) == "11,83%"


def test_parse_mes():
    assert parse_mes("jun/2026") == ("JUN", 2026)
    assert parse_mes("06/2026") == ("JUN", 2026)
    assert parse_mes(None) is None
    with pytest.raises(ValueError):
        parse_mes("junho")


def test_mes_do_arquivo():
    assert mes_do_arquivo("Apr_Shock Imports_v1_062026.xlsx") == ("JUN", 2026)
    assert mes_do_arquivo("Shock_Imports_Matriz_PME_2026_06_v1.xlsx") == ("JUN", 2026)
    assert mes_do_arquivo("qualquer.xlsx") is None


def test_achar_aba():
    abas = ["Relatório", "Base", "Dinamica", "Matriz  C COPART", "Matriz S COPART"]
    assert achar_aba(abas, "CC") == "Matriz  C COPART"
    assert achar_aba(abas, "SC") == "Matriz S COPART"
    assert achar_aba(["Base", "Matriz com coparticipação"], "SC") is None
    assert achar_aba(["Base", "Com coparticipação", "Sem coparticipação"], "SC") == "Sem coparticipação"
    assert norm("  Faixa   Etária ") == "faixa etaria"


def test_achar_abas_por_conteudo(tmp_path):
    """Nomes de aba fora do padrão: detecta pelo conteúdo (blocos 'OPERADORA')."""
    wb = openpyxl.Workbook()
    wb.active.title = "Base"
    for nome in ("Cotação Atual", "Cotação Sem Co-part"):
        ws = wb.create_sheet(nome)
        ws["B2"] = "OPERADORA"
        ws["C4"] = "VIDAS"
    from apresentador360.extract import achar_abas
    assert achar_abas(wb) == {"CC": "Cotação Atual", "SC": "Cotação Sem Co-part"}


# ---------- integração com os exemplos ------------------------------------ #
@precisa_amostras
def test_extracao_dinamica():
    m = carregar_matriz(MATRIZ)
    assert set(m.cenarios) == {"CC", "SC"}
    atual = m.cenarios["CC"].blocos[0]
    assert [p.vidas for p in atual.planos[:3]] == [1, 12, 3]
    assert atual.planos[3].nome is None            # "PLANO 4" é placeholder
    assert len(atual.faixas) == 10
    assert m.cenarios["CC"].reajuste_pct == pytest.approx(0.1183)
    assert m.demografia.titularidade == {"Funcionários": 10, "Dependentes": 6}
    assert m.demografia.genero == {"M": 5, "F": 11}


@precisa_amostras
def test_calculo_bate_com_excel_de_graficos():
    """Amil e Bradesco da Matriz existem no Excel de referência: os totais devem ser idênticos."""
    est = calcular(carregar_matriz(MATRIZ), [1, 3])
    cc = est.cenarios["CC"]
    assert cc.fatura_atual == pytest.approx(18219.6)
    assert cc.fatura_reajustada == pytest.approx(20374.97868)
    por_nome = {o.nome.split()[0]: o for o in cc.operadoras}
    assert por_nome["Amil"].total_mensal == pytest.approx(9201.52)
    assert por_nome["Amil"].reducao_pct == pytest.approx(-0.5483911838871186)
    assert por_nome["Bradesco"].total_mensal == pytest.approx(9346.47)
    sc = est.cenarios["SC"]
    assert {o.nome.split()[0]: o.total_mensal for o in sc.operadoras}["Amil"] == pytest.approx(12268.32)


@precisa_amostras
def test_ordem_do_menor_para_o_maior_custo():
    m = carregar_matriz(MATRIZ)
    est = calcular(m, [3, 1, 8])
    cc = est.cenarios["CC"].operadoras
    assert [o.total_mensal for o in cc] == sorted(o.total_mensal for o in cc)
    # o SC repete exatamente a mesma ordem
    assert [o.bloco_indice for o in est.cenarios["SC"].operadoras] == [o.bloco_indice for o in cc]


@precisa_amostras
def test_selecao_invalida():
    m = carregar_matriz(MATRIZ)
    with pytest.raises(ValueError):
        resolver_selecao(m, [1, 2, 3, 4, 5])      # mais de 4
    with pytest.raises(ValueError):
        resolver_selecao(m, [99])


@precisa_amostras
def test_pptx_gerado(tmp_path):
    saida = tmp_path / "saida.pptx"
    r = gerar(MATRIZ, PPTX, saida, operadoras=[8, 1, 3], mes_reajuste="SET/2026")
    prs = Presentation(str(saida))
    assert len(prs.slides) == 35
    textos = {i: "\n".join(sh.text_frame.text for sh in sl.shapes if sh.has_text_frame)
              for i, sl in enumerate(prs.slides, 1)}
    assert "Shock Imports" in textos[1]
    for esperado in ("R$18.220", "R$20.375", "SET", "11,83%", "R$9.202", "-55%", "R$9.346", "-54%"):
        assert esperado in textos[14], esperado
    for esperado in ("R$12.268", "-40%", "R$12.729", "-38%"):
        assert esperado in textos[18], esperado
    tabela15 = next(sh for sh in prs.slides[14].shapes if sh.has_table).table
    assert tabela15.cell(2, 4).text == "R$595,76"        # Amil, custo médio, plano 1
    assert tabela15.cell(13, 5).text == "R$2.045,72"     # Bradesco, total, plano 3
    tabela17 = next(sh for sh in prs.slides[16].shapes if sh.has_table).table
    assert tabela17.cell(0, 1).text == "R$ 247,96"       # Bradesco, 0 a 18, Efetivo E
    assert tabela17.cell(9, 3).text == "R$ 2.247,33"
    assert isinstance(r.avisos, list)


@precisa_amostras
def test_template_original_nao_e_alterado(tmp_path):
    antes = PPTX.read_bytes()
    gerar(MATRIZ, PPTX, tmp_path / "x.pptx", operadoras=[1])
    assert PPTX.read_bytes() == antes


@precisa_amostras
def test_menos_operadoras_remove_slides_sem_dados(tmp_path):
    gerar(MATRIZ, PPTX, tmp_path / "um.pptx", operadoras=[1])
    assert len(Presentation(str(tmp_path / "um.pptx")).slides) == 33   # sai 17 e 21


@precisa_amostras
def test_matriz_sem_sc(tmp_path):
    wb = openpyxl.load_workbook(MATRIZ, data_only=True)
    del wb["Matriz S COPART"]
    sem = tmp_path / "Matriz_2026_06_semSC.xlsx"
    wb.save(sem)
    r = gerar(sem, PPTX, tmp_path / "s.pptx", operadoras=[1, 3])
    assert list(r.estudo.cenarios) == ["CC"]
    assert len(Presentation(str(tmp_path / "s.pptx")).slides) == 30   # saem 18-21 (SC) e 17 (2 operadoras)


# ---------- robustez: a mesma Matriz "deformada" deve dar o mesmo resultado ---- #
def _assinatura(path):
    m = carregar_matriz(path)
    est = calcular(m, [1, 3, 8])
    return (
        {t: [(o.nome.split()[0], round(o.total_mensal, 2)) for o in c.operadoras]
         for t, c in est.cenarios.items()},
        m.demografia.titularidade, m.demografia.genero,
        {k: dict(v) for k, v in m.demografia.faixa_etaria.items()},
    )


def _variante(tmp_path, nome, mutar):
    wb = openpyxl.load_workbook(MATRIZ, data_only=True)
    mutar(wb)
    destino = tmp_path / f"{nome}.xlsx"
    wb.save(destino)
    return destino


def _renomear_abas(wb):
    wb["Matriz  C COPART"].title = "Cotação Com Copart"
    wb["Matriz S COPART"].title = "Cotação Sem Copart"
    wb["Base"].title = "Dados"


def _deslocar(wb):
    for n in ("Matriz  C COPART", "Matriz S COPART"):
        for faixa in list(wb[n].merged_cells.ranges):   # openpyxl não move mesclagens junto
            wb[n].unmerge_cells(str(faixa))
        wb[n].insert_rows(1, 4)
        wb[n].insert_cols(1, 3)


def _rotulos(wb):
    for n in ("Matriz  C COPART", "Matriz S COPART"):
        for row in wb[n].iter_rows():
            for c in row:
                v = norm(c.value)
                if v == "operadora":
                    c.value = "Operadora:"
                elif v == "vidas":
                    c.value = "Qtd. Vidas"
                elif v == "per capita":
                    c.value = "Valor"
    cab = {"Elegibilidade": "Titularidade", "Sexo": "Gênero", "Faixa etária": "Faixa Etária"}
    for c in wb["Base"][1]:
        c.value = cab.get(c.value, c.value)


def _sem_nome_de_cenario(wb):
    wb["Matriz  C COPART"].title = "Cenário 1"
    wb["Matriz S COPART"].title = "Cenário 2"


def _ordem_invertida(wb):
    _renomear_abas(wb)
    wb.move_sheet("Cotação Sem Copart", offset=-1)


@precisa_amostras
@pytest.mark.parametrize("mutar", [_renomear_abas, _deslocar, _rotulos, _sem_nome_de_cenario,
                                   _ordem_invertida],
                         ids=lambda f: f.__name__)
def test_leitura_resiste_a_variacoes(tmp_path, mutar):
    esperado = _assinatura(MATRIZ)
    assert _assinatura(_variante(tmp_path, mutar.__name__, mutar)) == esperado


def test_template_fora_do_padrao_da_erro_claro():
    from pptx import Presentation
    from apresentador360.slides import resumo_template, validar_template
    prs = Presentation()
    for _ in range(21):
        prs.slides.add_slide(prs.slide_layouts[6])
    problemas = validar_template(prs, ["CC"])
    assert any("slide 15" in p and "tabela" in p for p in problemas)
    assert "slide 15: 0 tabela(s)" in resumo_template(prs)
    assert validar_template(Presentation(), ["CC"])[0].startswith("o template tem 0 slides")


@precisa_amostras
def test_analise_financeira_continua_em_slide_duplicado(tmp_path):
    """Com 6 planos (a tabela comporta 4), o slide é duplicado e os planos 5 e 6 seguem no novo slide."""
    import copy as _copy
    from apresentador360.pptx_writer import remover_slide
    from apresentador360.slides import preencher_estudo

    m = carregar_matriz(MATRIZ)
    est = calcular(m, [1, 3])
    for cen in est.cenarios.values():
        base_a, base_o = cen.planos_atuais[:], [o.planos[:] for o in cen.operadoras]
        while len(cen.planos_atuais) < 6:
            i = len(cen.planos_atuais)
            novo = _copy.deepcopy(base_a[i % len(base_a)]); novo.nome = f"Atual extra {i + 1}"
            cen.planos_atuais.append(novo)
        for o, orig in zip(cen.operadoras, base_o):
            while len(o.planos) < 6:
                i = len(o.planos)
                p = _copy.deepcopy(orig[i % len(orig)]); p.nome = f"Plano extra {i + 1}"
                o.planos.append(p)

    prs = Presentation(str(PPTX))
    antes = len(prs.slides)
    avisos = []
    for s in preencher_estudo(prs, est, avisos):
        remover_slide(prs, s)
    assert len(prs.slides) == antes + 2                       # +1 em CC e +1 em SC
    tabelas = [next(sh for sh in s.shapes if sh.has_table).table for s in prs.slides
               if any(sh.has_table and len(sh.table.columns) == 7 and len(sh.table.rows) == 19 for sh in s.shapes)]
    assert len(tabelas) == 4                                  # 2 slides CC + 2 slides SC
    assert tabelas[0].cell(0, 1).text.strip() != ""
    assert tabelas[1].cell(0, 1).text.strip() == "Atual extra 5"
    assert tabelas[1].cell(5, 1).text.strip() == "Atual extra 6"
    assert tabelas[1].cell(10, 1).text.strip() == ""          # 3º grupo vazio na continuação
    assert any("duplicado" in a for a in avisos)
    prs.save(str(tmp_path / "x.pptx"))
    Presentation(str(tmp_path / "x.pptx"))                    # reabre sem erro


@precisa_amostras
def test_custo_por_faixa_cresce_com_mais_planos(tmp_path):
    """Com 4 operadoras de 6 planos (24 colunas), os slides de faixa passam de 2 para 3 por cenário."""
    import copy as _copy
    from apresentador360.pptx_writer import remover_slide
    from apresentador360.slides import preencher_estudo

    m = carregar_matriz(MATRIZ)
    est = calcular(m, [1, 3, 8, 6])
    for cen in est.cenarios.values():
        for o in cen.operadoras:
            orig = o.planos[:]
            while len(o.planos) < 6:
                p = _copy.deepcopy(orig[len(o.planos) % len(orig)]); p.nome = f"Extra {len(o.planos) + 1}"
                o.planos.append(p)
    prs = Presentation(str(PPTX))
    antes = len(prs.slides)
    avisos = []
    for s in preencher_estudo(prs, est, avisos):
        remover_slide(prs, s)
    tabelas = [sh.table for s in prs.slides for sh in s.shapes if sh.has_table and len(sh.table.columns) == 9]
    assert len(tabelas) >= 6                                   # 3 slides de faixa por cenário
    assert any("foram criados" in a for a in avisos)
    prs.save(str(tmp_path / "f.pptx")); Presentation(str(tmp_path / "f.pptx"))

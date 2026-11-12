"""Desconto da Matriz: o usuário escolhe usar o valor com ou sem desconto."""
import os
from pathlib import Path

import openpyxl
import pytest

from apresentador360.calc import calcular, listar_operadoras
from apresentador360.extract import carregar_matriz

PASTA = Path(os.environ.get("AP360_SAMPLES", "samples"))
PME = next(iter(sorted(PASTA.glob("Shock*Matriz*.xlsx"))), None) if PASTA.exists() else None
EMP = next(iter(sorted(PASTA.glob("*Corporate*.xlsx"))), None) if PASTA.exists() else None


@pytest.mark.skipif(PME is None, reason="amostra ausente")
def test_pme_desconto_no_quadro_final(tmp_path):
    wb = openpyxl.load_workbook(PME, data_only=True)
    ws = wb["Matriz  C COPART"]
    assert ws["C44"].value == "Bradesco"
    r = 44
    while ws.cell(r, 2).value != "Total / Fatura":
        r += 1
    fatura = None
    for c in range(3, 16):                                       # valor da fatura cheia no quadro final
        v = ws.cell(r, c).value
        if isinstance(v, (int, float)) and v > 0:
            fatura = v
            break
    assert fatura
    ws.cell(r + 1, 2).value = "Desc 10%"
    ws.cell(r + 1, 3).value = round(fatura * 0.9, 2)
    arq = tmp_path / "m.xlsx"
    wb.save(arq)

    m = carregar_matriz(arq)
    ops = {o["indice"]: o for o in listar_operadoras(m)}
    brad = next(i for i, o in ops.items() if o["nome"] == "Bradesco")
    outra = next(i for i, o in ops.items() if o["nome"] == "Amil Selecionada")
    assert ops[brad]["desconto"] is True and ops[outra]["desconto"] is False
    sem = calcular(m, [brad, outra], usar_desconto=False)
    com = calcular(m, [brad, outra], usar_desconto=True)
    nome = "Bradesco"
    s = next(o for o in sem.cenarios["CC"].operadoras if o.nome == nome)
    c = next(o for o in com.cenarios["CC"].operadoras if o.nome == nome)
    assert c.total_mensal == pytest.approx(s.total_mensal * 0.9, rel=1e-3)
    assert sum(p.total for p in c.planos) == pytest.approx(c.total_mensal)
    assert any("desconto" in a.lower() for a in com.avisos)
    outra_s = next(o for o in sem.cenarios["CC"].operadoras if o.nome != nome)
    outra_c = next(o for o in com.cenarios["CC"].operadoras if o.nome != nome)
    assert outra_s.total_mensal == outra_c.total_mensal           # sem desconto na Matriz: nada muda


@pytest.mark.skipif(PME is None, reason="amostra ausente")
def test_pme_sem_desconto_na_matriz_nao_muda():
    m = carregar_matriz(PME)
    assert not any(o["desconto"] for o in listar_operadoras(m))
    a, b = calcular(m, [1, 3], usar_desconto=True), calcular(m, [1, 3], usar_desconto=False)
    assert [o.total_mensal for o in a.cenarios["CC"].operadoras] == [o.total_mensal for o in b.cenarios["CC"].operadoras]


@pytest.mark.skipif(EMP is None, reason="amostra ausente")
def test_empresarial_escolhe_com_ou_sem_desconto():
    m = carregar_matriz(EMP, tipo="empresarial")
    ops = [o for o in listar_operadoras(m) if o["desconto"]]
    assert ops, "a amostra Vibropac traz operadora com desconto"
    idx = ops[0]["indice"]
    com = calcular(m, [idx], usar_desconto=True).cenarios["SC"].operadoras[0].total_mensal
    sem = calcular(m, [idx], usar_desconto=False).cenarios["SC"].operadoras[0].total_mensal
    assert com < sem


@pytest.mark.skipif(PME is None, reason="amostra ausente")
def test_pme_agravo_no_quadro_final(tmp_path):
    """"Agravo 8%": a fatura ajustada é MAIOR que a cheia e a escolha continua valendo."""
    wb = openpyxl.load_workbook(PME, data_only=True)
    ws = wb["Matriz  C COPART"]
    ws["B64"].value = "Agravo 8%"
    ws["C64"].value = round(ws["C63"].value * 1.08, 2)
    arq = tmp_path / "a.xlsx"
    wb.save(arq)
    m = carregar_matriz(arq)
    idx = next(o["indice"] for o in listar_operadoras(m) if o["nome"] == "Bradesco")
    com = calcular(m, [idx], usar_desconto=True).cenarios["CC"].operadoras[0]
    sem = calcular(m, [idx], usar_desconto=False).cenarios["CC"].operadoras[0]
    assert com.total_mensal == pytest.approx(sem.total_mensal * 1.08, rel=1e-3)
    assert any("agravo" in a.lower() for a in calcular(m, [idx]).avisos)

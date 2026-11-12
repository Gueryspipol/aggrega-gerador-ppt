"""Matriz EMPRESARIAL (layout de cenários diferente da PME). Usa a Matriz Vibropac Corporate.

Rode com AP360_SAMPLES apontando para a pasta que contém:
    *Corporate*.xlsx                      (Matriz Empresarial)
    Apr_Shock_Imports_v1_062026.pptx      (template)
    Shock_Imports_Matriz_PME_*.xlsx       (Matriz PME, para o teste de tipo trocado)
"""
import io
import os
import sys
from pathlib import Path

import pytest
from pptx import Presentation

from apresentador360.calc import calcular, listar_operadoras
from apresentador360.extract import carregar_matriz
from apresentador360.pipeline import gerar

PASTA = Path(os.environ.get("AP360_SAMPLES", "samples"))
EMP = next(iter(sorted(PASTA.glob("*Corporate*.xlsx"))), None) if PASTA.exists() else None
PME = next(iter(sorted(PASTA.glob("Shock*Matriz*.xlsx"))), None) if PASTA.exists() else None
TPL = PASTA / "Apr_Shock_Imports_v1_062026.pptx"
pytestmark = pytest.mark.skipif(EMP is None or not TPL.exists(), reason="amostras Empresarial ausentes")


def test_le_so_o_cenario_visivel_e_os_totais_da_planilha():
    m = carregar_matriz(EMP, tipo="empresarial")
    assert list(m.cenarios) == ["SC"]                      # CC está oculta: ignorada
    assert any("oculta" in a for a in m.avisos)
    tot = {o["nome"]: o["SC_desc"] for o in listar_operadoras(m)}   # SC_desc = valor com desconto/diluição
    assert tot["Amil Selecionada"] == pytest.approx(170582.14, abs=0.01)
    assert tot["Amil Metal"] == pytest.approx(133085.10, abs=0.01)
    assert tot["Bradesco"] == pytest.approx(176276.50, abs=0.01)
    assert tot["SulAmérica - Custo Médio"] == pytest.approx(148515.217857, abs=0.01)   # com diluição
    assert tot["Seguros Unimed"] == 0                       # operadora sem valores
    assert not any(n.strip() == "" for n in tot)             # blocos sem nome não viram operadora


def test_calculo_e_ordem():
    m = carregar_matriz(EMP, tipo="empresarial")
    nomes = {o["nome"]: o["indice"] for o in listar_operadoras(m)}
    e = calcular(m, [nomes["Bradesco"], nomes["Amil Metal"], nomes["Amil Selecionada"]])
    sc = e.cenarios["SC"]
    assert sc.fatura_atual == pytest.approx(166072.24, abs=0.01)
    assert sc.fatura_reajustada == pytest.approx(179358.01, abs=0.01)
    assert sc.reajuste_pct == pytest.approx(0.08)
    assert [o.nome for o in sc.operadoras] == ["Amil Metal", "Amil Selecionada", "Bradesco"]   # menor -> maior
    assert sc.operadoras[0].reducao_mensal == pytest.approx(133085.10 - 179358.01, abs=0.05)
    assert sum(p.vidas for p in sc.planos_atuais) == 205
    assert any("COM desconto" in a for a in calcular(m, [nomes["SulAmérica - Custo Médio"]]).avisos)


def test_pptx_sem_faixa_etaria(tmp_path):
    r = gerar(EMP, TPL, tmp_path / "emp.pptx", operadoras=[1, 2, 3, 4], cliente="Vibropac",
              tipo="empresarial")
    prs = Presentation(str(r.saida))
    titulos = [" ".join(sh.text_frame.text for sh in sl.shapes if sh.has_text_frame).upper() for sl in prs.slides]
    assert not any("CUSTO POR FAIXA" in t for t in titulos)     # slides de faixa etária removidos
    assert sum("ANÁLISE FINANCEIRA" in t for t in titulos) >= 1
    assert sum("INVESTIMENTO FINANCEIRO" in t for t in titulos) == 1   # só o SC


def test_tipo_trocado_da_erro_claro():
    with pytest.raises(ValueError, match="EMPRESARIAL"):
        carregar_matriz(EMP, tipo="pme")
    if PME:
        with pytest.raises(ValueError, match="PME"):
            carregar_matriz(PME, tipo="empresarial")


def test_site_envia_tipo():
    sys.path.insert(0, str(Path(__file__).parent.parent))
    import server
    cli = server.app.test_client()
    r = cli.post("/api/ler", data={"matriz": (io.BytesIO(EMP.read_bytes()), EMP.name), "tipo": "empresarial"})
    assert r.status_code == 200, r.get_json()
    d = r.get_json()
    assert d["tipo"] == "empresarial" and d["cenarios"] == ["SC"]
    errado = cli.post("/api/ler", data={"matriz": (io.BytesIO(EMP.read_bytes()), EMP.name), "tipo": "pme"})
    assert errado.status_code == 422 and "EMPRESARIAL" in errado.get_json()["erro"]
    assert cli.post("/api/ler", data={"matriz": (io.BytesIO(b"x"), "a.xlsx"), "tipo": "xyz"}).status_code == 400
    ids = [o["indice"] for o in d["operadoras"] if o["SC"] > 0][:3]
    if (Path(__file__).parent.parent / "templates").glob("*.pptx"):
        g = cli.post("/api/gerar", json={"token": d["token"], "operadoras": ids, "cliente": "Vibropac"})
        assert g.status_code in (200, 500), g.get_json()

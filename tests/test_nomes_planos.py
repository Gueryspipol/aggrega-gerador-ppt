"""Slides de análise financeira: nome do plano em negrito e só com a 1ª letra maiúscula."""
import os
from pathlib import Path

import pytest
from pptx import Presentation

from apresentador360.pipeline import gerar
from apresentador360.slides import formatar_nome_plano

PASTA = Path(os.environ.get("AP360_SAMPLES", "samples"))
PME = next(iter(sorted(PASTA.glob("Shock*Matriz*.xlsx"))), None) if PASTA.exists() else None
TPL = Path("templates/template.pptx")


@pytest.mark.parametrize("entrada,saida", [
    ("NACIONAL PLUS", "Nacional plus"), ("ESPECIAL MAIS", "Especial mais"), ("PREMIUM 900", "Premium 900"),
    ("P420", "P420"), ("NACIONAL III", "Nacional III"), ("Nacional E", "Nacional E"),
    ("Especial Mais", "Especial Mais"), ("", ""),
])
def test_formatar_nome(entrada, saida):
    assert formatar_nome_plano(entrada) == saida


@pytest.mark.skipif(PME is None or not TPL.exists(), reason="amostra/template ausente")
def test_nomes_em_negrito_na_financeira(tmp_path):
    prs = Presentation(str(gerar(PME, TPL, tmp_path / "x.pptx", operadoras=[1, 3, 8], cliente="X",
                                 tipo="pme").saida))
    achou = 0
    for sl in prs.slides:
        t = next((x.table for x in sl.shapes if x.has_table), None)
        if t is None or "PRODUTO" not in " ".join(c.text.upper() for r in t.rows for c in r.cells)[:400] and \
                "Produto" not in t.cell(0, 0).text:
            continue
        for g in range(len(t.rows) // 5):
            for c in range(1, len(t.columns)):
                cel = t.cell(5 * g, c)
                if cel.text.strip():
                    achou += 1
                    assert all(r.font.bold for r in cel.text_frame.paragraphs[0].runs), cel.text
    assert achou

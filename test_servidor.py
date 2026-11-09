"""Testa a API do site (Flask) com as amostras reais. Rode com AP360_SAMPLES apontando para a pasta."""
import io
import os
import sys
from pathlib import Path

import pytest

PASTA = Path(os.environ.get("AP360_SAMPLES", "samples"))
MATRIZ = next(iter(sorted(PASTA.glob("*Matriz*.xlsx"))), None) if PASTA.exists() else None
pytestmark = pytest.mark.skipif(MATRIZ is None or not (Path(__file__).parent.parent / "templates").glob("*.pptx"),
                                reason="amostras ou template ausentes")


def test_fluxo_completo():
    sys.path.insert(0, str(Path(__file__).parent.parent))
    import server
    cli = server.app.test_client()
    r = cli.post("/api/ler", data={"matriz": (io.BytesIO(MATRIZ.read_bytes()), MATRIZ.name)})
    assert r.status_code == 200
    d = r.get_json()
    assert len(d["operadoras"]) >= 4 and d["fatura_reajustada"] > d["fatura_atual"]
    ids = [o["indice"] for o in d["operadoras"][:4]]
    g = cli.post("/api/gerar", json={"token": d["token"], "operadoras": ids, "cliente": "Teste"})
    assert g.status_code == 200, g.get_json()
    assert cli.get(g.get_json()["download"]).status_code == 200
    assert cli.post("/api/gerar", json={"token": "naoexiste", "operadoras": ids}).status_code == 410


def test_arquivo_invalido():
    sys.path.insert(0, str(Path(__file__).parent.parent))
    import server
    r = server.app.test_client().post("/api/ler", data={"matriz": (io.BytesIO(b"x"), "a.txt")})
    assert r.status_code == 400

"""Catálogo de logos das operadoras.

Cada arquivo da pasta `logos/` vale para uma operadora e é reconhecido pelo NOME do arquivo:
`amil.png`, `bradesco.png`, `sulamerica.png`, `porto.png`... Uma operadora da Matriz usa o logo
cujo nome aparece dentro do nome dela ("Amil Metal" -> amil.png; "Porto linha PRO" -> porto.png;
"SulAmérica - adesão" -> sulamerica.png). Havendo mais de um, vale o nome de arquivo mais longo.
Para cadastrar uma operadora nova basta colocar o arquivo na pasta, sem mexer no código.
"""
from __future__ import annotations

import os
import re
import unicodedata
from pathlib import Path
from typing import Optional

EXTENSOES = (".png", ".jpg", ".jpeg")


def _norm(texto: str) -> str:
    t = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


def _compacto(texto: str) -> str:
    return _norm(texto).replace(" ", "")


def _chave_do_arquivo(caminho: Path) -> str:
    chave = _compacto(caminho.stem)
    chave = re.sub(r"^logo", "", chave)               # "logo_amil" -> "amil"
    return re.sub(r"\d+$", "", chave)                 # "amil1", "amil (2)" -> "amil"


def localizar_pasta_logos(pasta: Optional[str | Path] = None) -> Optional[Path]:
    candidatas = []
    if pasta:
        candidatas.append(Path(pasta))
    if os.environ.get("LOGOS_DIR"):
        candidatas.append(Path(os.environ["LOGOS_DIR"]))
    aqui = Path(__file__).resolve().parent
    candidatas += [aqui / "logos", aqui.parent / "logos", aqui.parent.parent / "logos",
                   Path.cwd() / "logos"]
    return next((p for p in candidatas if p.is_dir()), None)


# logos da própria Aggrega (não são de operadoras): logo.png, logo_branco.png, "logo (2).png"...
_LOGO_AGGREGA = re.compile(r"^logo([ _-]?(branco|aggrega))?( ?\(\d+\))?$", re.I)


def _pastas_de_logos(pasta: Optional[str | Path]) -> list[Path]:
    """Pasta `logos/` e, no repositório com arquivos soltos na raiz, a própria raiz."""
    pastas = []
    principal = localizar_pasta_logos(pasta)
    if principal:
        pastas.append(principal)
    aqui = Path(__file__).resolve().parent
    solto = (aqui / "pipeline.py").exists() and ((aqui / "server.py").exists() or (aqui / "app.py").exists())
    if solto and aqui not in pastas:
        pastas.append(aqui)
    return pastas


class Catalogo:
    def __init__(self, pasta: Optional[str | Path] = None):
        self.pastas = _pastas_de_logos(pasta)
        self.pasta = self.pastas[0] if self.pastas else None
        self.itens: list[tuple[str, Path]] = []
        vistos: set[str] = set()
        for pasta_logos in self.pastas:
            for arq in sorted(pasta_logos.iterdir()):
                if arq.suffix.lower() not in EXTENSOES or not arq.is_file():
                    continue
                if _LOGO_AGGREGA.match(arq.stem) or arq.stem.lower() in vistos:
                    continue
                chave = _chave_do_arquivo(arq)
                if len(chave) >= 3:
                    self.itens.append((chave, arq))
                    vistos.add(arq.stem.lower())

    def __bool__(self) -> bool:
        return bool(self.itens)

    def achar(self, nome: str) -> Optional[Path]:
        """Logo da operadora cujo nome de arquivo aparece em `nome` (o mais específico vence)."""
        alvo = _compacto(nome or "")
        melhores = [(len(ch), arq) for ch, arq in self.itens if ch in alvo]
        return max(melhores, key=lambda x: x[0])[1] if melhores else None

    def achar_todos(self, texto: str) -> list[Path]:
        """Todos os logos citados em `texto`, na ordem em que aparecem ("Bradesco e Amil")."""
        alvo = _compacto(texto or "")
        achados = {}
        for ch, arq in self.itens:
            pos = alvo.find(ch)
            if pos >= 0 and not any(ch != c2 and ch in c2 and c2 in alvo for c2, _ in self.itens):
                achados[arq] = pos
        return [a for a, _ in sorted(achados.items(), key=lambda x: x[1])]

    def nomes(self) -> list[str]:
        return [a.name for _, a in self.itens]

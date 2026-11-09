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


class Catalogo:
    def __init__(self, pasta: Optional[str | Path] = None):
        self.pasta = localizar_pasta_logos(pasta)
        self.itens: list[tuple[str, Path]] = []
        if self.pasta:
            for arq in sorted(self.pasta.iterdir()):
                if arq.suffix.lower() in EXTENSOES and arq.is_file():
                    chave = _chave_do_arquivo(arq)
                    if len(chave) >= 3:
                        self.itens.append((chave, arq))

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

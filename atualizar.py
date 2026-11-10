"""Atualização do projeto pelo próprio site.

Recebe arquivos (ou um .zip), confere o que mudou em relação ao repositório do GitHub e
publica tudo em um único commit. O Render detecta o commit e publica a nova versão sozinho.

Variáveis de ambiente:
  GITHUB_TOKEN   token fine-grained com "Contents: Read and write" só neste repositório
  GITHUB_REPO    dono/repositório (padrão: Gueryspipol/aggrega-gerador-ppt)
  GITHUB_BRANCH  ramo publicado pelo Render (padrão: main)
"""
from __future__ import annotations

import base64
import hashlib
import io
import json
import os
import posixpath
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass, field

REPO = os.environ.get("GITHUB_REPO", "Gueryspipol/aggrega-gerador-ppt")
RAMO = os.environ.get("GITHUB_BRANCH", "main")
API = "https://api.github.com"

EXTENSOES = {".py", ".html", ".css", ".js", ".png", ".jpg", ".jpeg", ".svg", ".ico", ".pptx",
             ".txt", ".md", ".toml", ".yaml", ".yml", ".json", ".cfg"}
IGNORAR_PARTES = {"__MACOSX", ".git", "__pycache__", ".pytest_cache", ".venv", "venv", "node_modules"}
IGNORAR_NOMES = {".DS_Store", "Thumbs.db", ".env"}
BLOQUEADOS = (".github/",)          # nunca altera automações do repositório
MAX_ARQUIVOS = 200
MAX_BYTES_ARQUIVO = 15 * 1024 * 1024


class ErroAtualizacao(Exception):
    """Erro com mensagem pronta para mostrar ao usuário."""

    def __init__(self, msg: str, codigo: int = 400):
        super().__init__(msg)
        self.codigo = codigo


@dataclass
class Plano:
    novos: list[str] = field(default_factory=list)
    alterados: list[str] = field(default_factory=list)
    iguais: list[str] = field(default_factory=list)
    ignorados: list[str] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)

    def resumo(self) -> dict:
        return {"novos": self.novos, "alterados": self.alterados, "iguais": self.iguais,
                "ignorados": self.ignorados, "avisos": self.avisos}


# ------------------------------------ GitHub ------------------------------------ #
def token() -> str:
    return os.environ.get("GITHUB_TOKEN", "").strip()


def _chamar(metodo: str, caminho: str, corpo: dict | None = None, *, tolera_404: bool = False):
    tk = token()
    if not tk:
        raise ErroAtualizacao("Falta a variável GITHUB_TOKEN no Render (Environment).", 503)
    dados = json.dumps(corpo).encode() if corpo is not None else None
    req = urllib.request.Request(
        f"{API}{caminho}", data=dados, method=metodo,
        headers={"Authorization": f"Bearer {tk}", "Accept": "application/vnd.github+json",
                 "X-GitHub-Api-Version": "2022-11-28", "User-Agent": "apresentador-360",
                 "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=40) as r:
            texto = r.read().decode() or "{}"
            return json.loads(texto)
    except urllib.error.HTTPError as e:
        if e.code == 404 and tolera_404:
            return None
        detalhe = ""
        try:
            detalhe = json.loads(e.read().decode()).get("message", "")
        except Exception:  # noqa: BLE001
            pass
        if e.code in (401, 403):
            raise ErroAtualizacao(
                "O GitHub recusou o token. Confira se ele não expirou e se tem permissão "
                f"“Contents: Read and write” no repositório {REPO}. ({detalhe})", 502)
        if e.code == 404:
            raise ErroAtualizacao(
                f"Repositório ou ramo não encontrado ({REPO}, {RAMO}). Confira o token. ({detalhe})", 502)
        raise ErroAtualizacao(f"Erro do GitHub ({e.code}): {detalhe}", 502)
    except (urllib.error.URLError, TimeoutError) as e:
        raise ErroAtualizacao(f"Não consegui falar com o GitHub agora: {e}. Tente de novo.", 502)


def ultimo_commit() -> dict | None:
    """Commit mais recente do ramo no GitHub (sha curto, mensagem, data)."""
    r = _chamar("GET", f"/repos/{REPO}/commits/{RAMO}")
    if not r:
        return None
    c = r.get("commit", {})
    return {"sha": r.get("sha", ""), "curto": r.get("sha", "")[:7],
            "mensagem": (c.get("message") or "").splitlines()[0] if c.get("message") else "",
            "data": (c.get("committer") or {}).get("date", "")}


def _sha_blob(dados: bytes) -> str:
    return hashlib.sha1(b"blob %d\0" % len(dados) + dados).hexdigest()  # noqa: S324 (formato do git)


def _arvore_atual() -> tuple[str, str, dict[str, str]]:
    """(sha do commit, sha da árvore, {caminho: sha do blob})"""
    ref = _chamar("GET", f"/repos/{REPO}/git/ref/heads/{RAMO}")
    sha_commit = ref["object"]["sha"]
    commit = _chamar("GET", f"/repos/{REPO}/git/commits/{sha_commit}")
    sha_arvore = commit["tree"]["sha"]
    arvore = _chamar("GET", f"/repos/{REPO}/git/trees/{sha_arvore}?recursive=1")
    arquivos = {i["path"]: i["sha"] for i in arvore.get("tree", []) if i.get("type") == "blob"}
    return sha_commit, sha_arvore, arquivos


# --------------------------------- entrada de arquivos --------------------------------- #
def _limpar_caminho(p: str) -> str | None:
    p = p.replace("\\", "/").strip().lstrip("/")
    p = posixpath.normpath(p) if p else ""
    if not p or p.startswith("..") or "/../" in f"/{p}/":
        return None
    return p


def _ler_entrada(nome: str, dados: bytes) -> dict[str, bytes]:
    """Transforma um upload em {caminho: bytes}; abre .zip."""
    if nome.lower().endswith(".zip"):
        try:
            zf = zipfile.ZipFile(io.BytesIO(dados))
        except zipfile.BadZipFile:
            raise ErroAtualizacao(f"O arquivo {nome} não é um .zip válido.")
        itens: dict[str, bytes] = {}
        total = 0
        for info in zf.infolist():
            if info.is_dir():
                continue
            total += info.file_size
            if total > 120 * 1024 * 1024:
                raise ErroAtualizacao("O .zip é grande demais.")
            itens[info.filename] = zf.read(info)
        # remove a pasta-raiz comum (ex.: "apresentador360/...")
        partes = [p.replace("\\", "/").split("/") for p in itens]
        if partes and all(len(p) > 1 for p in partes) and len({p[0] for p in partes}) == 1:
            itens = {"/".join(k.replace("\\", "/").split("/")[1:]): v for k, v in itens.items()}
        return itens
    return {nome: dados}


def montar_plano(entradas: list[tuple[str, bytes]]) -> tuple[Plano, dict[str, bytes], dict[str, str]]:
    """Compara os arquivos enviados com o repositório.

    Devolve (plano, {caminho_no_repo: bytes a gravar}, {caminho_no_repo: sha atual}).
    """
    brutos: dict[str, bytes] = {}
    for nome, dados in entradas:
        brutos.update(_ler_entrada(nome, dados))

    plano = Plano()
    _, _, repo = _arvore_atual()
    por_nome: dict[str, list[str]] = {}
    for caminho in repo:
        por_nome.setdefault(posixpath.basename(caminho), []).append(caminho)

    gravar: dict[str, bytes] = {}
    for original, dados in brutos.items():
        caminho = _limpar_caminho(original)
        base = posixpath.basename(caminho or "")
        partes = set((caminho or "").split("/"))
        if (caminho is None or base in IGNORAR_NOMES or partes & IGNORAR_PARTES
                or posixpath.splitext(base)[1].lower() not in EXTENSOES):
            plano.ignorados.append(original)
            continue
        if caminho.startswith(BLOQUEADOS):
            plano.ignorados.append(original)
            plano.avisos.append(f"{original}: arquivos de automação do GitHub não são alterados pelo site.")
            continue
        if len(dados) > MAX_BYTES_ARQUIVO:
            raise ErroAtualizacao(f"{original} tem mais de 15 MB.")
        # Arquivo solto que já existe em outra pasta do repositório: atualiza o existente.
        if "/" not in caminho and caminho not in repo and len(por_nome.get(caminho, [])) == 1:
            caminho = por_nome[caminho][0]
        if caminho.endswith(".py"):
            try:
                compile(dados, caminho, "exec")
            except SyntaxError as e:
                raise ErroAtualizacao(
                    f"{original} tem erro de sintaxe (linha {e.lineno}): {e.msg}. "
                    "Nada foi publicado.")
        gravar[caminho] = dados

    if len(gravar) > MAX_ARQUIVOS:
        raise ErroAtualizacao(f"Arquivos demais ({len(gravar)}). O limite é {MAX_ARQUIVOS}.")

    efetivos: dict[str, bytes] = {}
    for caminho, dados in sorted(gravar.items()):
        if caminho not in repo:
            plano.novos.append(caminho)
            efetivos[caminho] = dados
        elif repo[caminho] != _sha_blob(dados):
            plano.alterados.append(caminho)
            efetivos[caminho] = dados
        else:
            plano.iguais.append(caminho)

    # aviso de pacote incompleto: slides.py muda junto com pptx_writer.py / logos.py
    tocados = {posixpath.basename(c) for c in efetivos}
    if "slides.py" in tocados and not {"pptx_writer.py", "logos.py"} <= (
            tocados | {posixpath.basename(c) for c in repo}):
        plano.avisos.append("slides.py depende de pptx_writer.py e logos.py, que não estão no repositório.")
    return plano, efetivos, repo


def publicar(efetivos: dict[str, bytes], mensagem: str | None = None) -> dict:
    """Cria um único commit com todos os arquivos e avança o ramo."""
    if not efetivos:
        raise ErroAtualizacao("Nada para publicar: os arquivos já estão iguais aos do GitHub.")
    sha_commit, sha_arvore, _ = _arvore_atual()
    itens = []
    for caminho, dados in efetivos.items():
        blob = _chamar("POST", f"/repos/{REPO}/git/blobs",
                       {"content": base64.b64encode(dados).decode(), "encoding": "base64"})
        itens.append({"path": caminho, "mode": "100644", "type": "blob", "sha": blob["sha"]})
    arvore = _chamar("POST", f"/repos/{REPO}/git/trees", {"base_tree": sha_arvore, "tree": itens})
    nomes = ", ".join(posixpath.basename(c) for c in list(efetivos)[:6])
    mais = f" e mais {len(efetivos) - 6}" if len(efetivos) > 6 else ""
    msg = mensagem or f"Atualização pelo site: {nomes}{mais}"
    commit = _chamar("POST", f"/repos/{REPO}/git/commits",
                     {"message": msg, "tree": arvore["sha"], "parents": [sha_commit]})
    _chamar("PATCH", f"/repos/{REPO}/git/refs/heads/{RAMO}", {"sha": commit["sha"]})
    return {"sha": commit["sha"], "curto": commit["sha"][:7], "arquivos": sorted(efetivos)}

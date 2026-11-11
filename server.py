"""Apresentador 360 - servidor web (Flask).

Executar localmente:  python server.py   (abre em http://localhost:8000)
Em produção (Render):  gunicorn server:app --workers 1 --threads 4 --timeout 120
Senha opcional: defina a variável de ambiente APP_SENHA.
"""
from __future__ import annotations

import hashlib
import hmac
import os
import shutil
import sys
import tempfile
import time
import uuid
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory, session

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ / "src"))
sys.path.insert(0, str(RAIZ))

# Aceita o repositório organizado (src/apresentador360/) ou com os arquivos soltos na raiz.
if not ((RAIZ / "src" / "apresentador360" / "__init__.py").exists()
        or (RAIZ / "apresentador360" / "__init__.py").exists()):
    import importlib.util
    _spec = importlib.util.spec_from_file_location(
        "apresentador360", RAIZ / "__init__.py", submodule_search_locations=[str(RAIZ)])
    _pacote = importlib.util.module_from_spec(_spec)
    sys.modules["apresentador360"] = _pacote
    _spec.loader.exec_module(_pacote)

from apresentador360.calc import MAX_OPERADORAS, calcular, listar_operadoras  # noqa: E402
from apresentador360.extract import carregar_matriz, diagnosticar  # noqa: E402
from apresentador360.logos import Catalogo  # noqa: E402
from apresentador360.pipeline import gerar, mes_do_arquivo  # noqa: E402
import atualizar  # noqa: E402

VERSAO = "web 2"
SENHA = os.environ.get("APP_SENHA", "")
TEMPLATES = RAIZ / "templates"
PASTA_TRABALHO = Path(tempfile.gettempdir()) / "apresentador360"
VALIDADE_S = 2 * 3600

app = Flask(__name__, static_folder=None)
PASTAS_ESTATICAS = [RAIZ / "web" / "static", RAIZ / "assets", RAIZ]
PAGINAS = [RAIZ / "web" / "index.html", RAIZ / "index.html"]
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024
app.secret_key = os.environ.get("SECRET_KEY") or hashlib.sha256(
    (SENHA or "apresentador360").encode()).hexdigest()


# ----------------------------------- utilidades ----------------------------------- #
def _limpar() -> None:
    PASTA_TRABALHO.mkdir(parents=True, exist_ok=True)
    agora = time.time()
    for p in PASTA_TRABALHO.iterdir():
        try:
            if agora - p.stat().st_mtime > VALIDADE_S:
                shutil.rmtree(p, ignore_errors=True)
        except OSError:
            pass


def _template() -> Path | None:
    achados = sorted([*TEMPLATES.glob("*.pptx"), *RAIZ.glob("*.pptx")]) if TEMPLATES.exists() \
        else sorted(RAIZ.glob("*.pptx"))
    if not achados:
        return None
    return next((p for p in achados if "template" in p.name.lower()), achados[0])


def _pasta(token: str) -> Path:
    if not token.isalnum():
        raise ValueError("Sessão inválida")
    return PASTA_TRABALHO / token


def _erro(msg: str, codigo: int = 400, **extra):
    return jsonify({"erro": msg, **extra}), codigo


def _cenario_json(cen) -> dict:
    return {
        "tipo": cen.tipo, "reajuste_pct": cen.reajuste_pct, "fatura_atual": cen.fatura_atual,
        "fatura_reajustada": cen.fatura_reajustada, "total_vidas": cen.total_vidas,
        "operadoras": [{"nome": o.nome, "total": o.total_mensal, "reducao_mensal": o.reducao_mensal,
                        "reducao_pct": o.reducao_pct} for o in cen.operadoras],
    }


# ------------------------------------ acesso ------------------------------------ #
@app.before_request
def _proteger():
    if not SENHA:
        return None
    if request.path in ("/", "/api/entrar", "/api/estado") or request.path.startswith("/static/"):
        return None
    if session.get("ok"):
        return None
    return _erro("Acesso restrito. Entre com a senha.", 401)


@app.errorhandler(Exception)
def _erro_geral(e):
    """Qualquer falha vira uma resposta JSON legível (o site mostra a mensagem)."""
    from werkzeug.exceptions import HTTPException
    if isinstance(e, HTTPException):
        return _erro(f"{e.name}: {e.description}", e.code or 500)
    app.logger.exception("Erro não tratado")
    return _erro(f"Erro interno do servidor ({type(e).__name__}): {e}", 500)


@app.get("/")
def pagina():
    arq = next((p for p in PAGINAS if p.exists()), None)
    if arq is None:
        return "Página não encontrada: falta o arquivo index.html no repositório.", 500
    return send_from_directory(arq.parent, arq.name)


@app.get("/static/<nome>")
def estatico(nome: str):
    if not nome.lower().endswith((".png", ".jpg", ".jpeg", ".svg", ".ico")):
        return _erro("Arquivo não encontrado.", 404)
    for pasta in PASTAS_ESTATICAS:
        if (pasta / nome).is_file():
            return send_from_directory(pasta, nome)
    return _erro("Arquivo não encontrado.", 404)


@app.get("/api/estado")
def estado():
    return jsonify({"versao": VERSAO, "precisa_senha": bool(SENHA), "autenticado": bool(session.get("ok")) or not SENHA,
                    "max_operadoras": MAX_OPERADORAS, "template_ok": _template() is not None})


@app.post("/api/entrar")
def entrar():
    senha = (request.get_json(silent=True) or {}).get("senha", "")
    if SENHA and hmac.compare_digest(str(senha), SENHA):
        session["ok"] = True
        return jsonify({"ok": True})
    return _erro("Senha incorreta.", 401)


# --------------------------------- atualização do sistema --------------------------------- #
_PLANOS: dict[str, dict] = {}      # planos de atualização aguardando confirmação
_VALIDADE_PLANO_S = 15 * 60


def _exige_atualizacao():
    if not SENHA:
        return _erro("Por segurança, a aba Atualizar só funciona com a senha do site ativada "
                     "(variável APP_SENHA no Render).", 403)
    if not atualizar.token():
        return _erro("Falta a variável GITHUB_TOKEN no Render (Environment).", 503)
    return None


@app.get("/api/versao")
def versao():
    rodando = os.environ.get("RENDER_GIT_COMMIT", "")
    info = {"versao": VERSAO, "rodando": rodando[:7], "rodando_completo": rodando,
            "repositorio": atualizar.REPO, "ramo": atualizar.RAMO,
            "token_ok": bool(atualizar.token()), "senha_ok": bool(SENHA)}
    if atualizar.token():
        try:
            ult = atualizar.ultimo_commit()
            info["github"] = ult
            info["atualizado"] = bool(ult and rodando and ult["sha"] == rodando)
            info["publicando"] = bool(ult and rodando and ult["sha"] != rodando)
        except atualizar.ErroAtualizacao as e:
            info["aviso"] = str(e)
    return jsonify(info)


@app.post("/api/atualizar/analisar")
def atualizar_analisar():
    negado = _exige_atualizacao()
    if negado:
        return negado
    arquivos = [(f.filename or "", f.read()) for f in request.files.getlist("arquivos") if f.filename]
    if not arquivos:
        return _erro("Escolha os arquivos (ou um .zip) para atualizar.")
    try:
        plano, efetivos, _ = atualizar.montar_plano(arquivos)
    except atualizar.ErroAtualizacao as e:
        return _erro(str(e), e.codigo)
    agora = time.time()
    for k in [k for k, v in _PLANOS.items() if agora - v["quando"] > _VALIDADE_PLANO_S]:
        _PLANOS.pop(k, None)
    token = uuid.uuid4().hex
    _PLANOS[token] = {"efetivos": efetivos, "quando": agora}
    return jsonify({"token": token, **plano.resumo(), "total": len(efetivos)})


@app.post("/api/atualizar/publicar")
def atualizar_publicar():
    negado = _exige_atualizacao()
    if negado:
        return negado
    token = str((request.get_json(silent=True) or {}).get("token", ""))
    plano = _PLANOS.pop(token, None)
    if not plano:
        return _erro("Esta conferência expirou (o site pode ter reiniciado). Envie os arquivos de novo.", 410)
    try:
        resultado = atualizar.publicar(plano["efetivos"])
    except atualizar.ErroAtualizacao as e:
        return _erro(str(e), e.codigo)
    return jsonify(resultado)


TIPOS_MATRIZ = ("pme", "empresarial")


# ------------------------------------ fluxo ------------------------------------ #
@app.post("/api/ler")
def ler():
    _limpar()
    arq = request.files.get("matriz")
    if not arq or not arq.filename.lower().endswith(".xlsx"):
        return _erro("Envie um arquivo .xlsx.")
    tipo = (request.form.get("tipo") or "pme").strip().lower()
    if tipo not in TIPOS_MATRIZ:
        return _erro("Tipo de Matriz inválido (use PME ou Empresarial).")
    token = uuid.uuid4().hex
    pasta = _pasta(token)
    pasta.mkdir(parents=True, exist_ok=True)
    nome = Path(arq.filename).name
    caminho = pasta / nome
    arq.save(caminho)
    (pasta / "tipo.txt").write_text(tipo, encoding="utf-8")
    try:
        matriz = carregar_matriz(caminho, tipo=tipo)
        ops = [o for o in listar_operadoras(matriz) if any(o.get(t, 0) > 0 for t in matriz.cenarios)]
        if not ops:
            raise ValueError("Nenhuma operadora com valores foi encontrada na Matriz.")
        base = calcular(matriz, [ops[0]["indice"]])
    except Exception as e:  # noqa: BLE001
        try:
            diag = diagnosticar(caminho)
        except Exception:  # noqa: BLE001
            diag = ""
        shutil.rmtree(pasta, ignore_errors=True)
        return _erro(f"Não consegui ler a Matriz: {e}", 422, diagnostico=diag)

    ref = base.cenarios.get("CC") or next(iter(base.cenarios.values()))
    mes = mes_do_arquivo(caminho)
    catalogo = Catalogo()
    return jsonify({
        "token": token, "tipo": tipo, "arquivo": nome, "cliente": matriz.cliente or "",
        "mes_atual": f"{mes[0]}/{mes[1]}" if mes else "",
        "cenarios": list(matriz.cenarios),
        "vidas": ref.total_vidas, "fatura_atual": ref.fatura_atual,
        "fatura_reajustada": ref.fatura_reajustada, "reajuste_pct": ref.reajuste_pct,
        "operadora_atual": base.operadora_atual,
        "logo_atual": bool(catalogo.achar_todos(base.operadora_atual)),
        "logos_cadastrados": len(catalogo.itens),
        "operadoras": [{"indice": o["indice"], "nome": o["nome"] or "(sem nome)",
                        "logo": catalogo.achar(o["nome"] or "") is not None,
                        **{t: o.get(t, 0.0) for t in matriz.cenarios}} for o in ops],
        "diagnostico": diagnosticar(caminho),
    })


@app.post("/api/gerar")
def gerar_apresentacao():
    dados = request.get_json(silent=True) or {}
    try:
        pasta = _pasta(str(dados.get("token", "")))
        origem = next(p for p in pasta.glob("*.xlsx"))
    except (ValueError, StopIteration):
        return _erro("A sessão expirou. Envie a Matriz novamente.", 410)
    tpl = _template()
    if tpl is None:
        return _erro("Template .pptx não encontrado no repositório (pasta templates/).", 500)
    try:
        tipo = (pasta / "tipo.txt").read_text(encoding="utf-8").strip()
    except OSError:
        tipo = "pme"
    escolha = dados.get("operadoras") or []
    cliente = (dados.get("cliente") or "").strip()
    try:
        saida_dir = pasta / "saida"
        saida_dir.mkdir(exist_ok=True)
        base = "".join(ch for ch in (cliente or "apresentacao") if ch.isalnum() or ch in " _-").strip() or "apresentacao"
        saida = saida_dir / f"{base}.pptx"
        r = gerar(origem, tpl, saida, operadoras=[int(i) for i in escolha], cliente=cliente or None,
                  mes_atual=(dados.get("mes_atual") or None), mes_reajuste=(dados.get("mes_reajuste") or None),
                  operadora_atual=(dados.get("operadora_atual") or None),
                  tipo=tipo)
    except Exception as e:  # noqa: BLE001
        return _erro(f"Não foi possível gerar: {e}", 422)
    return jsonify({
        "arquivo": saida.name, "avisos": r.avisos,
        "cenarios": {t: _cenario_json(c) for t, c in r.estudo.cenarios.items()},
        "download": f"/api/baixar/{pasta.name}/{saida.name}",
    })


@app.get("/api/baixar/<token>/<nome>")
def baixar(token: str, nome: str):
    try:
        caminho = _pasta(token) / "saida" / Path(nome).name
    except ValueError:
        return _erro("Sessão inválida.", 400)
    if not caminho.exists():
        return _erro("Arquivo não encontrado. Gere a apresentação novamente.", 404)
    return send_file(caminho, as_attachment=True, download_name=caminho.name,
                     mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), debug=False)

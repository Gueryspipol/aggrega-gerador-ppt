"""Interface Streamlit do Apresentador 360.

Executar:  streamlit run app.py
No Streamlit Community Cloud, o "Main file path" é:  app.py
"""
import sys
import tempfile
from pathlib import Path

import streamlit as st

RAIZ = Path(__file__).parent
sys.path.insert(0, str(RAIZ / "src"))
sys.path.insert(0, str(RAIZ))          # aceita também apresentador360/ direto na raiz

_na_pasta = ((RAIZ / "src" / "apresentador360" / "__init__.py").exists()
             or (RAIZ / "apresentador360" / "__init__.py").exists())
if not _na_pasta and (RAIZ / "__init__.py").exists() and (RAIZ / "pipeline.py").exists():
    # arquivos soltos na raiz do repositório (sem a pasta): a própria raiz vira o pacote
    import importlib.util
    _spec = importlib.util.spec_from_file_location(
        "apresentador360", RAIZ / "__init__.py", submodule_search_locations=[str(RAIZ)])
    _pacote = importlib.util.module_from_spec(_spec)
    sys.modules["apresentador360"] = _pacote
    _spec.loader.exec_module(_pacote)
    _na_pasta = True

if not _na_pasta:
    st.error("Pasta do robô não encontrada no repositório.")
    st.markdown(
        "O repositório precisa ter **`src/apresentador360/`** com estes 9 arquivos: "
        "`__init__.py`, `__main__.py`, `calc.py`, `cli.py`, `extract.py`, `formatos.py`, "
        "`models.py`, `pipeline.py`, `pptx_writer.py` e `slides.py`.")
    st.write("Arquivos encontrados na raiz:", sorted(p.name for p in RAIZ.iterdir()))
    st.stop()

import base64  # noqa: E402

from apresentador360 import formatos as f  # noqa: E402
from apresentador360.calc import MAX_OPERADORAS, calcular, listar_operadoras  # noqa: E402
from apresentador360.extract import carregar_matriz, diagnosticar  # noqa: E402
from apresentador360.pipeline import gerar, mes_do_arquivo  # noqa: E402

VERSAO = "v4"
TEMPLATES = RAIZ / "templates"
ASSETS = RAIZ / "assets"

st.set_page_config(page_title="Apresentador 360 · Aggrega", page_icon="📊", layout="wide")

st.markdown("""
<style>
#MainMenu, footer {visibility: hidden;}
.block-container {padding-top: 1.2rem; max-width: 1180px;}
.topo {background: linear-gradient(120deg, #14264F 0%, #1F4FD1 100%); border-radius: 18px;
       padding: 26px 32px; display: flex; align-items: center; justify-content: space-between;
       margin-bottom: 18px;}
.topo h1 {color: #fff; font-size: 1.9rem; margin: 0; padding: 0; line-height: 1.15;}
.topo p {color: #CFE0FF; margin: 4px 0 0 0; font-size: 0.98rem;}
.topo img {height: 54px;}
.passo {display: flex; align-items: center; gap: 12px; margin: 22px 0 8px 0;}
.passo .n {background: #1F4FD1; color: #fff; width: 32px; height: 32px; border-radius: 50%;
           display: flex; align-items: center; justify-content: center; font-weight: 700;}
.passo .t {font-size: 1.25rem; font-weight: 700; color: #14264F;}
.passo .d {color: #6B7A99; font-size: 0.92rem; margin-left: 4px;}
.chip {display: inline-block; padding: 2px 10px; border-radius: 999px; font-size: 0.78rem;
       font-weight: 600; margin-left: 6px;}
.chip.verde {background: #DDF5E6; color: #17794A;}
.chip.azul {background: #E1EAFF; color: #1F4FD1;}
.chip.cinza {background: #ECEFF6; color: #5B6988;}
div[data-testid="stVerticalBlockBorderWrapper"] {border-radius: 14px; background: #fff;}
div.stButton > button[kind="primary"], div.stDownloadButton > button {
    border-radius: 10px; padding: 0.65rem 1.6rem; font-weight: 700;}
</style>
""", unsafe_allow_html=True)


def _img64(nome: str) -> str:
    caminho = ASSETS / nome
    return base64.b64encode(caminho.read_bytes()).decode() if caminho.exists() else ""


_logo = _img64("logo_branco.png")
st.markdown(
    f"""<div class="topo"><div><h1>Apresentador 360</h1>
    <p>Da Matriz de cotações para a apresentação PowerPoint, no padrão Aggrega.</p></div>
    {'<img src="data:image/png;base64,' + _logo + '">' if _logo else ''}</div>""",
    unsafe_allow_html=True)


def passo(n: int, titulo: str, dica: str = "") -> None:
    st.markdown(f'<div class="passo"><div class="n">{n}</div><div class="t">{titulo}</div>'
                f'<div class="d">{dica}</div></div>', unsafe_allow_html=True)


# ----------------------------- template fixo (sem upload) ----------------------------- #
_candidatos = sorted([*TEMPLATES.glob("*.pptx"), *RAIZ.glob("*.pptx")]) if TEMPLATES.exists() \
    else sorted(RAIZ.glob("*.pptx"))
if not _candidatos:
    st.error("Template não encontrado. Suba o arquivo .pptx padrão na pasta `templates/` do "
             "repositório no GitHub e reinicie o app.")
    st.stop()
tpl_fixo = next((p for p in _candidatos if "template" in p.name.lower()), _candidatos[0])


@st.cache_data(show_spinner="Lendo a Matriz…")
def _ler(nome: str, dados: bytes):
    pasta = Path(tempfile.mkdtemp())
    caminho = pasta / nome
    caminho.write_bytes(dados)
    matriz = carregar_matriz(caminho)
    ops = [o for o in listar_operadoras(matriz) if any(o.get(t, 0) > 0 for t in matriz.cenarios)]
    base = None
    if ops:
        try:
            base = calcular(matriz, [ops[0]["indice"]])
        except Exception:  # noqa: BLE001
            base = None
    return caminho, matriz, ops, base, diagnosticar(caminho)


# ----------------------------------- passo 1 ----------------------------------- #
passo(1, "Envie a Matriz", "arquivo .xlsx com as abas de cotação")
matriz_up = st.file_uploader("Matriz (.xlsx)", type=["xlsx"], label_visibility="collapsed")

if matriz_up is None:
    st.info("Envie a Matriz para começar. O robô reconhece sozinho as operadoras e os cenários "
            "com e sem coparticipação.")
    st.caption(f"Apresentador 360 · robô {VERSAO}")
    st.stop()

try:
    matriz_path, matriz, ops, base, diag = _ler(matriz_up.name, matriz_up.getvalue())
except Exception as e:  # noqa: BLE001
    st.error("Não consegui ler a Matriz. Abaixo está o que o robô enxergou no arquivo "
             "(me envie este texto para eu ajustar a leitura):")
    try:
        _tmp = Path(tempfile.mkdtemp()) / matriz_up.name
        _tmp.write_bytes(matriz_up.getvalue())
        _diag = diagnosticar(_tmp)
    except Exception:  # noqa: BLE001
        _diag = ""
    st.code(f"{e}\n\n{_diag}" if "Abas do arquivo" not in str(e) else str(e))
    st.stop()

cenarios = list(matriz.cenarios)
cc = base.cenarios.get("CC") or next(iter(base.cenarios.values())) if base else None
with st.container(border=True):
    c = st.columns(5)
    c[0].metric("Cenários", " + ".join(cenarios))
    c[1].metric("Operadoras na Matriz", len(ops))
    if cc:
        c[2].metric("Vidas", cc.total_vidas)
        c[3].metric("Fatura atual", f.brl_inteiro(cc.fatura_atual))
        c[4].metric("Fatura reajustada", f.brl_inteiro(cc.fatura_reajustada),
                    f"{cc.reajuste_pct:.2%}".replace(".", ","), delta_color="off")
    with st.expander("O que o robô reconheceu na Matriz"):
        st.code(diag)

# ----------------------------------- passo 2 ----------------------------------- #
passo(2, "Escolha as operadoras", f"até {MAX_OPERADORAS} entre as {len(ops)} encontradas · "
      "entram na apresentação do menor para o maior custo")

chave = f"{matriz_up.name}-{matriz_up.size}"
ref = "CC" if "CC" in cenarios else cenarios[0]
ordenadas = sorted(ops, key=lambda o: (o.get(ref, 0) or float("inf"), o["indice"]))
marcadas = sum(bool(st.session_state.get(f"op_{chave}_{o['indice']}")) for o in ordenadas)
fatura_reaj = cc.fatura_reajustada if cc else 0

with st.container(border=True):
    cab = st.columns([0.6, 3.2] + [1.6] * len(cenarios) + [1.4])
    cab[1].markdown("**Operadora**")
    for j, t in enumerate(cenarios):
        cab[2 + j].markdown(f"**{t}**")
    cab[-1].markdown(f"**{ref} vs reajuste**")
    for pos, o in enumerate(ordenadas):
        k = f"op_{chave}_{o['indice']}"
        linha = st.columns([0.6, 3.2] + [1.6] * len(cenarios) + [1.4])
        linha[0].checkbox(o["nome"] or "(sem nome)", key=k, label_visibility="collapsed",
                          disabled=(marcadas >= MAX_OPERADORAS and not st.session_state.get(k)))
        selo = ' <span class="chip verde">menor custo</span>' if pos == 0 else ""
        linha[1].markdown(f"{o['nome'] or '(sem nome)'}{selo}", unsafe_allow_html=True)
        for j, t in enumerate(cenarios):
            linha[2 + j].write(f.brl_centavos(o[t]))
        if fatura_reaj and o.get(ref):
            linha[-1].write(f.pct_inteiro(o[ref] / fatura_reaj - 1))

escolha = [o["indice"] for o in ordenadas if st.session_state.get(f"op_{chave}_{o['indice']}")]
if escolha:
    nomes = {o["indice"]: o["nome"] for o in ops}
    st.caption(f"{len(escolha)} de {MAX_OPERADORAS} selecionadas: " + " → ".join(nomes[i] for i in escolha))
else:
    st.caption(f"Marque de 1 a {MAX_OPERADORAS} operadoras.")

# ----------------------------------- passo 3 ----------------------------------- #
passo(3, "Dados do cliente", "aparecem na capa e no resumo")
col1, col2, col3 = st.columns(3)
cliente = col1.text_input("Cliente", matriz.cliente or "")
mes = mes_do_arquivo(matriz_path)
mes_atual = col2.text_input("Mês atual", f"{mes[0]}/{mes[1]}" if mes else "", placeholder="JUN/2026")
mes_reaj = col3.text_input("Mês do reajuste", "", placeholder="SET/2026")

# ----------------------------------- passo 4 ----------------------------------- #
passo(4, "Gere a apresentação")
if st.button("Gerar apresentação", type="primary", disabled=not escolha):
    try:
        with st.status("Montando a apresentação…", expanded=True) as estado:
            st.write("Calculando os cenários e a ordem de menor custo")
            saida = Path(tempfile.mkdtemp()) / f"{cliente or 'apresentacao'}.pptx"
            st.write("Preenchendo o template (slides 1, 13 a 21)")
            r = gerar(matriz_path, tpl_fixo, saida, operadoras=escolha, cliente=cliente or None,
                      mes_atual=mes_atual or None, mes_reajuste=mes_reaj or None)
            estado.update(label="Apresentação pronta", state="complete", expanded=False)
    except Exception as e:  # noqa: BLE001
        st.error(f"Erro ao gerar ({type(e).__name__}): {e or 'sem mensagem'}")
        st.exception(e)
        try:
            from pptx import Presentation
            from apresentador360.slides import resumo_template
            st.code(resumo_template(Presentation(str(tpl_fixo))))
        except Exception:  # noqa: BLE001
            pass
    else:
        st.session_state["resultado"] = (saida.read_bytes(), saida.name, r.estudo, list(r.avisos))

resultado = st.session_state.get("resultado")
if resultado:
    dados, nome_arq, estudo, avisos = resultado
    st.success("Apresentação gerada. Confira o resumo e baixe o arquivo.")
    st.download_button("⬇ Baixar PowerPoint", dados, file_name=nome_arq,
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    abas = st.tabs([f"{'Com' if t == 'CC' else 'Sem'} coparticipação ({t})" for t in estudo.cenarios])
    for aba, (tipo, cen) in zip(abas, estudo.cenarios.items()):
        with aba:
            cols = st.columns(len(cen.operadoras) + 2)
            cols[0].metric("Fatura atual", f.brl_inteiro(cen.fatura_atual))
            cols[1].metric("Reajustada", f.brl_inteiro(cen.fatura_reajustada), f.pct_2casas(cen.reajuste_pct),
                           delta_color="off")
            for j, o in enumerate(cen.operadoras):
                cols[2 + j].metric(o.nome, f.brl_inteiro(o.total_mensal), f.pct_inteiro(o.reducao_pct),
                                   delta_color="inverse")
    if avisos:
        with st.expander(f"Conferir antes de enviar ({len(avisos)})"):
            for a in avisos:
                st.markdown(f"- {a}")

st.caption(f"Apresentador 360 · robô {VERSAO}")

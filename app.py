import streamlit as st
import openpyxl
from io import BytesIO
from datetime import date, datetime

st.set_page_config(page_title="PRISMA", layout="wide", page_icon="🤝",
                   initial_sidebar_state="expanded")

TEMPLATE_PATH = "template/Prisma - Template.xlsx"
TAB_NAMES = ["Cadastro", "Endereços", "Contatos", "Perfil",
             "Comercial", "Regras", "Licenças", "Estratégico", "Crédito"]

if "_tab" not in st.session_state:
    st.session_state["_tab"] = 0

# ══════════════════════════════════════════════════════════════════════════════
# CSS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
#MainMenu, footer, header, [data-testid="stDecoration"], [data-testid="stToolbar"] { display:none !important; }

.stApp, [data-testid="stAppViewContainer"], [data-testid="stMain"],
[data-testid="stMainBlockContainer"], .main .block-container {
    background-color: #0f0f0e !important;
}

/* sidebar */
[data-testid="stSidebar"] {
    background-color: #141312 !important;
    border-right: 1px solid rgba(255,255,255,0.06) !important;
}
[data-testid="stSidebar"] section { padding: 0 20px !important; }

/* text */
p, span, div, li, a, [data-testid="stMarkdownContainer"] p { color: #f8f8f7 !important; }
h1, h2, h3, h4 { color: #f8f8f7 !important; font-family: 'Space Grotesk', sans-serif !important; }

/* labels */
label, [data-testid="stWidgetLabel"] > div,
[data-testid="stWidgetLabel"] p {
    color: #766f6b !important;
    font-size: 0.68rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
}

/* inputs */
input[type="text"], input[type="number"], input[type="email"],
input[type="date"], textarea {
    background-color: #1e1d1b !important;
    border: 1px solid rgba(255,255,255,0.09) !important;
    border-radius: 8px !important;
    color: #f8f8f7 !important;
    font-size: 0.86rem !important;
    font-family: 'Inter', sans-serif !important;
    transition: border-color .15s ease !important;
}
input:focus, textarea:focus {
    border-color: rgba(244,62,1,.45) !important;
    box-shadow: 0 0 0 3px rgba(244,62,1,.08) !important;
    outline: none !important;
}

/* selectbox */
[data-baseweb="select"] > div {
    background-color: #1e1d1b !important;
    border-color: rgba(255,255,255,0.09) !important;
    border-radius: 8px !important;
    color: #f8f8f7 !important;
}
[data-baseweb="select"] span, [data-baseweb="select"] p { color: #f8f8f7 !important; font-size:.86rem !important; }
[data-baseweb="popover"], [data-baseweb="popover"] ul, [data-baseweb="popover"] [role="listbox"] {
    background-color: #201f1d !important;
    border: 1px solid rgba(255,255,255,0.1) !important;
    border-radius: 10px !important;
}
[data-baseweb="popover"] [role="option"] { color: #f8f8f7 !important; }
[data-baseweb="popover"] [role="option"]:hover { background: rgba(244,62,1,.12) !important; }
[data-baseweb="popover"] [aria-selected="true"] { background: rgba(244,62,1,.18) !important; }

/* date */
[data-testid="stDateInputField"] input { color: #f8f8f7 !important; }

/* ── RADIO como tab bar ── */
[data-testid="stRadio"] { margin-bottom: 28px !important; }
[data-testid="stRadio"] > label { display: none !important; }
[data-testid="stRadio"] [role="radiogroup"] {
    display: flex !important;
    gap: 0 !important;
    border-bottom: 1px solid rgba(255,255,255,0.07) !important;
    flex-wrap: wrap !important;
}
[data-testid="stRadio"] [role="radiogroup"] label {
    padding: 10px 16px !important;
    margin: 0 !important;
    cursor: pointer !important;
    font-size: .72rem !important;
    font-weight: 500 !important;
    letter-spacing: .08em !important;
    text-transform: uppercase !important;
    color: #5a5450 !important;
    border-bottom: 2px solid transparent !important;
    white-space: nowrap !important;
    position: relative !important;
    bottom: -1px !important;
    transition: color .15s ease !important;
    display: flex !important;
    align-items: center !important;
    gap: 0 !important;
    background: transparent !important;
}
[data-testid="stRadio"] [role="radiogroup"] label:hover { color: #a09890 !important; }
[data-testid="stRadio"] input[type="radio"] { display: none !important; }
[data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {
    color: #f8f8f7 !important;
    border-bottom-color: #f43e01 !important;
}
[data-testid="stRadio"] [role="radiogroup"] label > div { display: none !important; }
[data-testid="stRadio"] [role="radiogroup"] label p {
    font-size: .72rem !important;
    font-weight: 500 !important;
    letter-spacing: .08em !important;
    text-transform: uppercase !important;
    margin: 0 !important;
}

/* buttons */
.stButton > button {
    background: transparent !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    color: #f8f8f7 !important;
    border-radius: 9999px !important;
    font-size: .75rem !important;
    font-weight: 500 !important;
    letter-spacing: .08em !important;
    text-transform: uppercase !important;
    padding: 8px 20px !important;
    transition: all .15s ease !important;
}
.stButton > button:hover { border-color: rgba(255,255,255,.3) !important; background: rgba(255,255,255,.04) !important; }
.stButton > button[kind="primary"] {
    background: #f43e01 !important; border-color: #f43e01 !important;
    color: #fff !important; font-weight: 600 !important;
}
.stButton > button[kind="primary"]:hover { background: #d93500 !important; border-color: #d93500 !important; }
.stDownloadButton > button {
    background: rgba(16,230,141,.08) !important;
    border: 1px solid rgba(16,230,141,.25) !important;
    color: #10e68d !important; border-radius: 9999px !important;
    font-size:.75rem !important; font-weight:600 !important;
    letter-spacing:.08em !important; text-transform:uppercase !important;
    width:100% !important;
}
.stDownloadButton > button:hover { background: rgba(16,230,141,.14) !important; border-color: rgba(16,230,141,.45) !important; }

hr { border-color: rgba(255,255,255,0.06) !important; margin: 16px 0 !important; }

[data-testid="stAlert"] { background: rgba(244,62,1,.08) !important; border: 1px solid rgba(244,62,1,.2) !important; border-radius:8px !important; color:#f8f8f7 !important; }

::-webkit-scrollbar { width:4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: rgba(255,255,255,.1); border-radius:4px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════
def fmt_date(d):
    if not d: return ""
    if isinstance(d, (date, datetime)): return d.strftime("%d/%m/%Y")
    return str(d)

def gs(key, default=""):
    return st.session_state.get(key, default)

def gs_date(key):
    return fmt_date(st.session_state.get(key))

def fill_template(data: dict) -> bytes:
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    for key, value in data.items():
                        token = "{{" + key + "}}"
                        if token in cell.value:
                            cell.value = cell.value.replace(token, str(value) if value else "")
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def sec(label, title, desc=""):
    st.markdown(f"""
    <div style="margin:4px 0 28px">
      <div style="font-size:.6rem;font-weight:600;letter-spacing:.22em;text-transform:uppercase;color:#766f6b;margin-bottom:10px">{label}</div>
      <div style="font-family:'Space Grotesk',sans-serif;font-size:1.35rem;font-weight:600;color:#f8f8f7;letter-spacing:-.015em;line-height:1.2">{title}</div>
      {"" if not desc else f'<div style="font-size:.8rem;color:#766f6b;margin-top:6px;line-height:1.6">{desc}</div>'}
      <div style="width:100%;height:1px;background:rgba(255,255,255,.06);margin-top:20px"></div>
    </div>
    """, unsafe_allow_html=True)

def sub(text):
    st.markdown(f'<div style="font-size:.6rem;font-weight:600;letter-spacing:.18em;text-transform:uppercase;color:#5a5450;margin:22px 0 10px">{text}</div>', unsafe_allow_html=True)

def col_head(*labels, cols):
    for col, lbl in zip(cols, labels):
        col.markdown(f"<div style='font-size:.6rem;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:#5a5450;padding-bottom:6px'>{lbl}</div>", unsafe_allow_html=True)

def sim_nao(label, key, lv="visible"):
    return st.selectbox(label, ["", "Sim", "Não"], key=key, label_visibility=lv)

def nivel(label, key, lv="visible"):
    return st.selectbox(label, ["", "1 — Baixo", "2 — Médio", "3 — Alto"], key=key, label_visibility=lv)

def s_doc(label, key, lv="visible"):
    return st.selectbox(label, ["", "Pendente", "Entregue", "Não se aplica"], key=key, label_visibility=lv)

def s_lic(label, key, lv="visible"):
    return st.selectbox(label, ["", "Válida", "Vencida", "Pendente", "Não se aplica"], key=key, label_visibility=lv)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style="padding:32px 0 24px">
      <div style="font-family:'Space Grotesk',sans-serif;font-size:1.5rem;font-weight:700;letter-spacing:.12em;color:#f8f8f7">PRISMA</div>
      <div style="font-size:.58rem;color:#5a5450;letter-spacing:.14em;text-transform:uppercase;margin-top:5px">Ficha Técnica do Cliente</div>
    </div>
    """, unsafe_allow_html=True)
    st.divider()
    for i, name in enumerate(TAB_NAMES):
        active = st.session_state["_tab"] == i
        color = "#f8f8f7" if active else "#766f6b"
        dot   = f'<span style="color:#f43e01;font-size:.5rem;margin-right:6px">●</span>' if active else ""
        st.markdown(
            f'<div style="display:flex;align-items:center;gap:8px;padding:5px 0;cursor:pointer">'
            f'<span style="font-family:IBM Plex Mono,monospace;font-size:.55rem;color:#3a3430">0{i+1}</span>'
            f'{dot}<span style="font-size:.78rem;color:{color};font-weight:{"500" if active else "400"}">{name}</span>'
            f'</div>', unsafe_allow_html=True)
    st.divider()
    st.markdown('<div style="font-size:.6rem;color:#3a3430;letter-spacing:.06em">BIOTROP · uso interno</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB BAR (radio estilizado)
# ══════════════════════════════════════════════════════════════════════════════
selected = st.radio("nav", TAB_NAMES, index=st.session_state["_tab"],
                    horizontal=True, label_visibility="collapsed", key="_nav_radio")
new_tab = TAB_NAMES.index(selected)
if new_tab != st.session_state["_tab"]:
    st.session_state["_tab"] = new_tab
    st.rerun()

ct = st.session_state["_tab"]

# ══════════════════════════════════════════════════════════════════════════════
# CONTEÚDO
# ══════════════════════════════════════════════════════════════════════════════

# ── 0 · CADASTRO ─────────────────────────────────────────────────────────────
if ct == 0:
    sec("Bloco 01", "Dados Cadastrais", "Identificação legal e fiscal do cliente.")
    c1, c2, c3 = st.columns([3, 3, 2])
    with c1:
        st.text_input("Razão Social", key="razao_social")
        st.text_input("CNPJ", placeholder="00.000.000/0000-00", key="cnpj")
    with c2:
        st.text_input("Nome Fantasia", key="nome_fantasia")
        st.text_input("CNAE", placeholder="0000-0/00", key="cnae")
    with c3:
        st.text_input("Inscrição Estadual", key="inscricao_estadual")
        st.text_input("Inscrição Municipal", key="inscricao_municipal")
    c4, c5, c6 = st.columns(3)
    with c4: st.date_input("Data de Abertura", value=None, key="data_abertura")
    with c5: st.text_input("Grupo Econômico", key="grupo_economico")
    with c6: st.date_input("Cliente Desde", value=None, key="data_inicio_relacionamento")
    sub("Responsável Biotrop")
    rb1, rb2, rb3 = st.columns(3)
    with rb1: st.text_input("Nome completo", key="ctrl_responsavel_biotrop")
    with rb2: st.text_input("E-mail", key="ctrl_responsavel_biotrop_email")
    with rb3: st.text_input("Telefone", key="ctrl_responsavel_biotrop_tel")

# ── 1 · ENDEREÇOS ─────────────────────────────────────────────────────────────
elif ct == 1:
    sec("Bloco 02", "Endereços", "Endereço fiscal e pontos de entrega.")
    sub("Endereço Fiscal")
    e1, e2, e3 = st.columns([5, 1, 2])
    with e1: st.text_input("Logradouro", key="end_fiscal_logradouro")
    with e2: st.text_input("Nº", key="end_fiscal_numero")
    with e3: st.text_input("Complemento", key="end_fiscal_complemento")
    e4, e5, e6, e7 = st.columns([3, 3, 1, 2])
    with e4: st.text_input("Bairro", key="end_fiscal_bairro")
    with e5: st.text_input("Município", key="end_fiscal_municipio")
    with e6: st.text_input("UF", max_chars=2, key="end_fiscal_estado")
    with e7: st.text_input("CEP", placeholder="00000-000", key="end_fiscal_cep")
    sub("Endereços de Entrega")
    h1, h2, h3, h4 = st.columns([1, 2, 3, 4])
    col_head("", "Identificação", "Município / Estado", "Observações", cols=[h1, h2, h3, h4])
    for n in range(1, 5):
        ec1, ec2, ec3, ec4 = st.columns([1, 2, 3, 4])
        with ec1:
            st.markdown(f"<div style='padding-top:10px;font-family:IBM Plex Mono,monospace;font-size:.75rem;color:#5a5450'>0{n}</div>", unsafe_allow_html=True)
        with ec2: st.text_input("", key=f"end_entrega_{n}_id", label_visibility="collapsed", placeholder="Nome da unidade")
        with ec3: st.text_input("", key=f"end_entrega_{n}_municipio_estado", label_visibility="collapsed", placeholder="Cidade / UF")
        with ec4: st.text_input("", key=f"end_entrega_{n}_obs", label_visibility="collapsed", placeholder="—")

# ── 2 · CONTATOS ──────────────────────────────────────────────────────────────
elif ct == 2:
    sec("Bloco 03", "Contatos", "Dois contatos por área: comercial, financeiro, técnico e logística.")
    grupos = [("Comercial","com"), ("Financeiro","fin"), ("Resp. Técnico","tec"), ("Logística","log")]
    for g_label, prefix in grupos:
        sub(g_label)
        h1, h2, h3, h4, h5 = st.columns([1, 3, 2, 2, 3])
        col_head("", "Nome", "Cargo", "Tel / Whatsapp", "E-mail", cols=[h1, h2, h3, h4, h5])
        for idx in (1, 2):
            c1, c2, c3, c4, c5 = st.columns([1, 3, 2, 2, 3])
            with c1: st.markdown(f"<div style='padding-top:10px;font-family:IBM Plex Mono,monospace;font-size:.7rem;color:#5a5450'>{idx}°</div>", unsafe_allow_html=True)
            with c2: st.text_input("", key=f"{prefix}{idx}_nome",  label_visibility="collapsed", placeholder="—")
            with c3: st.text_input("", key=f"{prefix}{idx}_cargo", label_visibility="collapsed", placeholder="—")
            with c4: st.text_input("", key=f"{prefix}{idx}_tel",   label_visibility="collapsed", placeholder="—")
            with c5: st.text_input("", key=f"{prefix}{idx}_email", label_visibility="collapsed", placeholder="—")

# ── 3 · PERFIL ────────────────────────────────────────────────────────────────
elif ct == 3:
    sec("Bloco 04", "Perfil do Cliente", "Caracterização do negócio e relação com os produtos Biotrop.")
    p1, p2, p3 = st.columns([2, 2, 3])
    with p1:
        st.selectbox("Tipo de cliente", ["","Distribuidor","Revendedor","Cooperativa","Produtor Rural","Importador","Outro"], key="tipo_cliente")
        st.text_input("Área Total (ha)", key="area_total_ha")
    with p2:
        st.text_input("Culturas Principais", key="culturas_principais")
        st.text_input("Região de Atuação", key="regiao_atuacao")
    with p3:
        st.text_area("Produtos Biotrop que já compra", height=112, key="produtos_utilizados", placeholder="Liste os produtos já utilizados…")

# ── 4 · COMERCIAL ─────────────────────────────────────────────────────────────
elif ct == 4:
    sec("Bloco 05", "Condições Comerciais e Financeiras", "Pagamento, crédito e política de bonificação.")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.text_input("Condição de Pagamento", placeholder="ex: 30/60/90 DDL", key="condicao_pagamento")
        st.text_input("Prazo Médio (dias)", key="prazo_medio_dias")
    with c2:
        st.text_input("Limite de Crédito (R$)", key="limite_credito")
        st.selectbox("Forma de Pagamento", ["","Boleto","PIX","Transferência Bancária","Cheque","Cartão","Misto"], key="forma_pagamento")
    with c3:
        st.text_area("Política de Bonificação", height=112, key="politica_bonificacao")
    st.text_area("Condições Comerciais Especiais", height=72, key="condicoes_especiais", placeholder="—")

# ── 5 · REGRAS ────────────────────────────────────────────────────────────────
elif ct == 5:
    sec("Bloco 06", "Regras de Faturamento e Logística", "Requisitos operacionais para emissão de nota e entrega.")
    sub("Faturamento")
    f1, f2, f3, f4 = st.columns([2, 1, 1, 1])
    with f1: st.text_input("Data limite para faturamento", placeholder="ex: dia 20", key="fat_data_limite")
    with f2: sim_nao("Exige PO?",          "fat_exige_po")
    with f3: sim_nao("Exige Contrato?",     "fat_exige_contrato")
    with f4: sim_nao("Conferência de NF?",  "fat_conferencia_nf")
    f5, f6, f7 = st.columns([2, 2, 3])
    with f5: st.text_input("Prazo de Envio para NF", placeholder="ex: 48h antes", key="fat_prazo_nf")
    with f6: st.text_input("Shelf life mínimo", placeholder="ex: 12 meses", key="fat_shelf_life")
    with f7: st.text_area("Observações de Faturamento", height=72, key="fat_observacoes", placeholder="—")
    sub("Logística")
    l1, l2, l3, l4, l5 = st.columns([2, 1, 2, 1, 2])
    with l1: st.selectbox("Tipo de entrega", ["","CIF","FOB","CIF e FOB"], key="log_tipo_entrega")
    with l2: sim_nao("Agendamento?",        "log_agendamento")
    with l3: st.text_input("Prazo mínimo agendamento", placeholder="ex: 48h", key="log_prazo_agendamento")
    with l4: sim_nao("Aviso?",              "log_aviso_entrega")
    with l5: st.text_input("Antecedência aviso (h)", key="log_antecedencia_aviso")
    l6, l7, l8, l9 = st.columns([2, 2, 1, 1])
    with l6: st.text_input("Dias de recebimento", placeholder="ex: Seg a Sex", key="log_dias_recebimento")
    with l7: st.text_input("Horário de recebimento", placeholder="ex: 07h–17h", key="log_horario_recebimento")
    with l8: sim_nao("NF antecipada?",      "log_nf_antecipada")
    with l9: st.text_input("Antecedência NF (h)", key="log_antecedencia_nf")
    l10, l11, l12 = st.columns([1, 2, 4])
    with l10: sim_nao("Romaneio?",          "log_romaneio")
    with l11: st.text_input("Restrição transportadora", key="log_restricao_transportadora")
    with l12: st.text_area("Regras de acesso / EPI / Observações", height=72, key="log_acesso_obs", placeholder="—")

# ── 6 · LICENÇAS ──────────────────────────────────────────────────────────────
elif ct == 6:
    sec("Bloco 07", "Licenças e Documentação", "Status de validade das licenças regulatórias do cliente.")
    licencas = [("Licença Ambiental","lic_ambiental"),("Licença de Operação","lic_operacao"),
                ("Registro MAPA","lic_mapa"),("Alvará Municipal","lic_alvara"),("Certificação ISO","lic_iso")]
    hd = st.columns([3, 2, 2, 4])
    col_head("Documento", "Status", "Validade", "Observações", cols=hd)
    for nome_lic, prefix in licencas:
        c1, c2, c3, c4 = st.columns([3, 2, 2, 4])
        with c1: st.markdown(f"<div style='padding-top:10px;font-size:.85rem;color:#a09890'>{nome_lic}</div>", unsafe_allow_html=True)
        with c2: s_lic("", f"{prefix}_status", lv="collapsed")
        with c3: st.date_input("", value=None, key=f"{prefix}_validade", label_visibility="collapsed")
        with c4: st.text_input("", key=f"{prefix}_obs", label_visibility="collapsed", placeholder="—")

# ── 7 · ESTRATÉGICO ───────────────────────────────────────────────────────────
elif ct == 7:
    sec("Bloco 08", "Informações Estratégicas e Complexidade", "Inteligência competitiva e avaliação operacional.")
    sub("Informações Estratégicas")
    es1, es2, es3 = st.columns(3)
    with es1:
        st.text_area("Concorrentes utilizados", height=100, key="estrat_concorrentes", placeholder="—")
        st.text_input("Potencial de volume estimado", placeholder="ex: 500 t/ano", key="estrat_potencial_volume")
    with es2:
        st.text_area("Histórico de relacionamentos", height=100, key="estrat_historico", placeholder="—")
        st.text_input("Participação estimada (%)", placeholder="ex: 30%", key="estrat_participacao")
    with es3:
        st.selectbox("Classificação de risco", ["","Baixo","Médio","Alto"], key="estrat_risco")
        st.text_area("Observações estratégicas", height=100, key="estrat_observacoes", placeholder="—")
    sub("Complexidade Operacional")
    co1, co2 = st.columns([2, 5])
    with co1: st.selectbox("Classificação Geral", ["","Baixa","Média","Alta"], key="compl_classificacao")
    with co2: st.text_area("Justificativa", height=72, key="compl_justificativa", placeholder="—")
    hd = st.columns([4, 1, 4])
    col_head("Critério", "Nível (1–3)", "Observação", cols=hd)
    for nome_crit, prefix in [("Exigências logísticas","compl_log"),("Burocracia de faturamento","compl_fat"),
                               ("Dificuldade de acesso","compl_acesso"),("Nível de exigência operacional","compl_operacional")]:
        cr1, cr2, cr3 = st.columns([4, 1, 4])
        with cr1: st.markdown(f"<div style='padding-top:10px;font-size:.85rem;color:#a09890'>{nome_crit}</div>", unsafe_allow_html=True)
        with cr2: nivel("", f"{prefix}_nivel", lv="collapsed")
        with cr3: st.text_input("", key=f"{prefix}_obs", label_visibility="collapsed", placeholder="—")

# ── 8 · CRÉDITO + CONTROLE ────────────────────────────────────────────────────
elif ct == 8:
    def render_docs(prefix, bloco, titulo, docs):
        sec(f"Bloco {bloco}", titulo)
        hd = st.columns([5, 2, 2, 3])
        col_head("Documento", "Status", "Data de entrega", "Observações", cols=hd)
        for n, nome_doc in enumerate(docs, start=1):
            c1, c2, c3, c4 = st.columns([5, 2, 2, 3])
            with c1: st.markdown(f"<div style='padding-top:10px;font-size:.82rem;color:#a09890'>{nome_doc}</div>", unsafe_allow_html=True)
            with c2: s_doc("", f"{prefix}_doc{n}_status", lv="collapsed")
            with c3: st.date_input("", value=None, key=f"{prefix}_doc{n}_data", label_visibility="collapsed")
            with c4: st.text_input("", key=f"{prefix}_doc{n}_obs", label_visibility="collapsed", placeholder="—")

    render_docs("cred_ltda", "09", "Análise de Crédito — Empresas LTDA", [
        "Documentos pessoais dos sócios e cônjuges (RG / CPF)",
        "Comprovante de endereço (sócios)",
        "Certidão de casamento dos sócios (se aplicável)",
        "Contrato social e última alteração contratual",
        "Balanço Patrimonial e DRE — últimos 2 anos (assinados)",
        "Imposto de Renda dos sócios — último exercício",
    ])
    render_docs("cred_coop", "10", "Análise de Crédito — Cooperativas / S.A. / Usinas", [
        "Documentos pessoais dos dirigentes",
        "Comprovante de endereço dos dirigentes",
        "Certidão de casamento dos sócios (se aplicável)",
        "Estatuto Social",
        "Ata de Eleição da Diretoria vigente",
        "Balanço Patrimonial e DRE — últimos 2 anos (assinados)",
    ])
    sec("Controle", "Controle Interno")
    ci1, ci2, ci3 = st.columns(3)
    with ci1:
        st.text_input("Cadastro realizado por", key="ctrl_cadastrado_por")
        st.date_input("Data do cadastro", value=date.today(), key="ctrl_data_cadastro")
    with ci2:
        st.text_input("Responsável pela atualização", key="ctrl_responsavel_atualizacao")
        st.date_input("Última atualização", value=date.today(), key="ctrl_ultima_atualizacao")
    with ci3:
        st.selectbox("Status do cadastro", ["","Em preenchimento","Completo","Em revisão","Aprovado"], key="ctrl_status")

# ══════════════════════════════════════════════════════════════════════════════
# NAVEGAÇÃO
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("<div style='margin-top:40px;border-top:1px solid rgba(255,255,255,0.06);padding-top:24px'></div>",
            unsafe_allow_html=True)

nav_l, nav_space, nav_r = st.columns([2, 4, 2])

with nav_l:
    if ct > 0:
        if st.button(f"← {TAB_NAMES[ct - 1]}", key="_nav_prev"):
            st.session_state["_tab"] = ct - 1
            st.rerun()

with nav_r:
    if ct < len(TAB_NAMES) - 1:
        if st.button(f"Próximo: {TAB_NAMES[ct + 1]} →", key="_nav_next"):
            st.session_state["_tab"] = ct + 1
            st.rerun()
    else:
        if st.button("Gerar Ficha", key="_nav_gerar", type="primary"):
            st.session_state["_gerar"] = True

# download após geração
if "arquivo_gerado" in st.session_state:
    st.markdown("""
    <div style="text-align:center;margin-top:20px">
      <div style="font-size:.72rem;color:#10e68d;letter-spacing:.06em;text-transform:uppercase;margin-bottom:12px">✓ Ficha gerada com sucesso</div>
    </div>
    """, unsafe_allow_html=True)
    dl_l, dl_c, dl_r = st.columns([2, 3, 2])
    with dl_c:
        st.download_button("↓  Baixar Excel",
            data=st.session_state["arquivo_gerado"],
            file_name=st.session_state.get("_nome_arquivo", "PRISMA.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

st.markdown("<div style='margin-bottom:60px'></div>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# LÓGICA — Gerar Ficha
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.get("_gerar"):
    st.session_state["_gerar"] = False
    razao = gs("razao_social").strip()
    if not razao:
        st.error("Preencha a Razão Social (aba Cadastro) antes de gerar a ficha.")
    else:
        log_obs = gs("log_acesso_obs")
        lic_pfx  = ["lic_ambiental","lic_operacao","lic_mapa","lic_alvara","lic_iso"]
        compl_cx = ["log","fat","acesso","operacional"]
        data = {
            "razao_social": razao,
            "nome_fantasia": gs("nome_fantasia"), "cnpj": gs("cnpj"),
            "inscricao_estadual": gs("inscricao_estadual"), "inscricao_municipal": gs("inscricao_municipal"),
            "cnae": gs("cnae"), "data_abertura": gs_date("data_abertura"),
            "grupo_economico": gs("grupo_economico"),
            "data_inicio_relacionamento": gs_date("data_inicio_relacionamento"),
            "ctrl_responsavel_biotrop":          gs("ctrl_responsavel_biotrop"),
            "ctrl_responsavel_biotrop-email":    gs("ctrl_responsavel_biotrop_email"),
            "ctrl_responsavel_biotrop-telefone": gs("ctrl_responsavel_biotrop_tel"),
            "end_fiscal_logradouro": gs("end_fiscal_logradouro"), "end_fiscal_numero": gs("end_fiscal_numero"),
            "end_fiscal_complemento": gs("end_fiscal_complemento"), "end_fiscal_bairro": gs("end_fiscal_bairro"),
            "end_fiscal_municipio": gs("end_fiscal_municipio"), "end_fiscal_estado": gs("end_fiscal_estado"),
            "end_fiscal_cep": gs("end_fiscal_cep"),
            **{f"end_entrega_{n}_{c}": gs(f"end_entrega_{n}_{c}") for n in range(1,5) for c in ("id","municipio_estado","obs")},
            **{f"{p}{i}_{f}": gs(f"{p}{i}_{f}") for p in ("com","fin","tec","log") for i in (1,2) for f in ("nome","cargo","tel","email")},
            "tipo_cliente": gs("tipo_cliente"), "culturas_principais": gs("culturas_principais"),
            "area_total_ha": gs("area_total_ha"), "regiao_atuacao": gs("regiao_atuacao"),
            "produtos_utilizados": gs("produtos_utilizados"),
            "condicao_pagamento": gs("condicao_pagamento"), "prazo_medio_dias": gs("prazo_medio_dias"),
            "limite_credito": gs("limite_credito"), "forma_pagamento": gs("forma_pagamento"),
            "politica_bonificacao": gs("politica_bonificacao"), "condicoes_especiais": gs("condicoes_especiais"),
            "fat_data_limite": gs("fat_data_limite"), "fat_exige_po": gs("fat_exige_po"),
            "fat_exige_contrato": gs("fat_exige_contrato"), "fat_conferencia_nf": gs("fat_conferencia_nf"),
            "fat_prazo_nf": gs("fat_prazo_nf"), "fat_shelf_life": gs("fat_shelf_life"),
            "fat_observacoes": gs("fat_observacoes"),
            "log_tipo_entrega": gs("log_tipo_entrega"), "log_agendamento": gs("log_agendamento"),
            "log_prazo_agendamento": gs("log_prazo_agendamento"), "log_dias_recebimento": gs("log_dias_recebimento"),
            "log_horario_recebimento": gs("log_horario_recebimento"), "log_aviso_entrega": gs("log_aviso_entrega"),
            "log_antecedencia_aviso": gs("log_antecedencia_aviso"), "log_nf_antecipada": gs("log_nf_antecipada"),
            "log_antecedencia_nf": gs("log_antecedencia_nf"), "log_romaneio": gs("log_romaneio"),
            "log_restricao_transportadora": gs("log_restricao_transportadora"),
            "log_regras_acesso": log_obs, "log_observacoes": log_obs,
            **{f"{p}_status":   gs(f"{p}_status")         for p in lic_pfx},
            **{f"{p}_validade": gs_date(f"{p}_validade")   for p in lic_pfx},
            **{f"{p}_obs":      gs(f"{p}_obs")             for p in lic_pfx},
            "estrat_concorrentes": gs("estrat_concorrentes"), "estrat_potencial_volume": gs("estrat_potencial_volume"),
            "estrat_participacao": gs("estrat_participacao"), "estrat_historico": gs("estrat_historico"),
            "estrat_risco": gs("estrat_risco"), "estrat_observacoes": gs("estrat_observacoes"),
            "compl_classificacao": gs("compl_classificacao"), "compl_justificativa": gs("compl_justificativa"),
            **{f"compl_{x}_nivel": gs(f"compl_{x}_nivel") for x in compl_cx},
            **{f"compl_{x}_obs":   gs(f"compl_{x}_obs")   for x in compl_cx},
            **{f"cred_ltda_doc{n}_status": gs(f"cred_ltda_doc{n}_status")         for n in range(1,7)},
            **{f"cred_ltda_doc{n}_data":   gs_date(f"cred_ltda_doc{n}_data")      for n in range(1,7)},
            **{f"cred_ltda_doc{n}_obs":    gs(f"cred_ltda_doc{n}_obs")            for n in range(1,7)},
            **{f"cred_coop_doc{n}_status": gs(f"cred_coop_doc{n}_status")         for n in range(1,7)},
            **{f"cred_coop_doc{n}_data":   gs_date(f"cred_coop_doc{n}_data")      for n in range(1,7)},
            **{f"cred_coop_doc{n}_obs":    gs(f"cred_coop_doc{n}_obs")            for n in range(1,7)},
            "ctrl_cadastrado_por": gs("ctrl_cadastrado_por"),
            "ctrl_data_cadastro": gs_date("ctrl_data_cadastro"),
            "ctrl_ultima_atualizacao": gs_date("ctrl_ultima_atualizacao"),
            "ctrl_responsavel_atualizacao": gs("ctrl_responsavel_atualizacao"),
            "ctrl_status": gs("ctrl_status"),
        }
        try:
            st.session_state["arquivo_gerado"] = fill_template(data)
            st.session_state["_nome_arquivo"]  = f"PRISMA_{razao.replace(' ','_')}.xlsx"
            st.rerun()
        except FileNotFoundError:
            st.error("Template não encontrado em `template/Prisma - Template.xlsx`.")
        except Exception as e:
            st.error(f"Erro ao gerar a ficha: {e}")

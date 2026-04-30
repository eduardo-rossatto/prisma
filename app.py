import streamlit as st
import openpyxl
from io import BytesIO
from datetime import date, datetime

st.set_page_config(
    page_title="PRISMA",
    layout="wide",
    page_icon="🤝",
    initial_sidebar_state="expanded",
)

TEMPLATE_PATH = "template/Prisma - Template.xlsx"

# ══════════════════════════════════════════════════════════════════════════════
# CSS — Groq-inspired dark theme
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

/* ── reset ── */
html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }
#MainMenu, footer, header, [data-testid="stDecoration"], [data-testid="stToolbar"] { display: none !important; }

/* ── backgrounds ── */
.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stMain"],
[data-testid="stMainBlockContainer"],
.main .block-container {
    background-color: #0f0f0e !important;
}

/* ── sidebar ── */
[data-testid="stSidebar"] {
    background-color: #141312 !important;
    border-right: 1px solid rgba(255,255,255,0.06) !important;
}
[data-testid="stSidebar"] section { padding: 0 20px !important; }

/* ── text ── */
p, span, div, li, a, [data-testid="stMarkdownContainer"] p { color: #f8f8f7 !important; }
h1, h2, h3, h4 { color: #f8f8f7 !important; font-family: 'Space Grotesk', sans-serif !important; }

/* ── labels ── */
label,
[data-testid="stWidgetLabel"] > div,
[data-testid="stWidgetLabel"] p,
[data-testid="stTextInput"] label,
[data-testid="stSelectbox"] label,
[data-testid="stTextArea"] label,
[data-testid="stDateInput"] label,
[data-testid="stNumberInput"] label {
    color: #766f6b !important;
    font-size: 0.68rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
}

/* ── inputs ── */
input[type="text"], input[type="number"], input[type="email"],
input[type="date"], input[type="time"], textarea {
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

/* CNPJ / CEP / CNAE — monospace */
[data-testid="stTextInput"]:has(input[aria-label*="CNPJ"]) input,
[data-testid="stTextInput"]:has(input[aria-label*="CEP"]) input,
[data-testid="stTextInput"]:has(input[aria-label*="CNAE"]) input {
    font-family: 'IBM Plex Mono', monospace !important;
    letter-spacing: .05em !important;
}

/* ── selectbox ── */
[data-baseweb="select"] > div {
    background-color: #1e1d1b !important;
    border-color: rgba(255,255,255,0.09) !important;
    border-radius: 8px !important;
    color: #f8f8f7 !important;
}
[data-baseweb="select"] [data-testid="stMarkdownContainer"] p,
[data-baseweb="select"] span { color: #f8f8f7 !important; font-size: .86rem !important; }
[data-baseweb="popover"],
[data-baseweb="popover"] ul,
[data-baseweb="popover"] [role="listbox"] {
    background-color: #201f1d !important;
    border: 1px solid rgba(255,255,255,0.1) !important;
    border-radius: 10px !important;
}
[data-baseweb="popover"] [role="option"] { color: #f8f8f7 !important; }
[data-baseweb="popover"] [role="option"]:hover { background-color: rgba(244,62,1,.12) !important; }
[data-baseweb="popover"] [aria-selected="true"] { background-color: rgba(244,62,1,.18) !important; }

/* ── date input ── */
[data-testid="stDateInputField"] input { color: #f8f8f7 !important; }

/* ── tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: transparent !important;
    border-bottom: 1px solid rgba(255,255,255,0.07) !important;
    gap: 0 !important;
    padding: 0 !important;
    margin-bottom: 28px !important;
}
.stTabs [data-baseweb="tab"] {
    background: transparent !important;
    color: #5a5450 !important;
    font-size: 0.75rem !important;
    font-weight: 500 !important;
    letter-spacing: .06em !important;
    text-transform: uppercase !important;
    border-radius: 0 !important;
    border: none !important;
    border-bottom: 2px solid transparent !important;
    padding: 10px 18px !important;
    margin-bottom: -1px !important;
    transition: color .15s ease !important;
}
.stTabs [aria-selected="true"] {
    color: #f8f8f7 !important;
    border-bottom-color: #f43e01 !important;
}
.stTabs [data-baseweb="tab"]:hover { color: #a09890 !important; background: transparent !important; }
.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"] { display: none !important; }
.stTabs [data-baseweb="tab-panel"] { padding: 0 !important; }

/* ── buttons ── */
.stButton > button {
    background: transparent !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    color: #f8f8f7 !important;
    border-radius: 9999px !important;
    font-size: 0.75rem !important;
    font-weight: 500 !important;
    letter-spacing: .08em !important;
    text-transform: uppercase !important;
    padding: 8px 20px !important;
    transition: all .15s ease !important;
}
.stButton > button:hover {
    border-color: rgba(255,255,255,.3) !important;
    background: rgba(255,255,255,.04) !important;
}
.stButton > button[kind="primary"] {
    background: #f43e01 !important;
    border-color: #f43e01 !important;
    color: #fff !important;
    font-weight: 600 !important;
}
.stButton > button[kind="primary"]:hover {
    background: #d93500 !important;
    border-color: #d93500 !important;
}
.stDownloadButton > button {
    background: rgba(16,230,141,.08) !important;
    border: 1px solid rgba(16,230,141,.25) !important;
    color: #10e68d !important;
    border-radius: 9999px !important;
    font-size: .75rem !important;
    font-weight: 600 !important;
    letter-spacing: .08em !important;
    text-transform: uppercase !important;
    width: 100% !important;
    transition: all .15s ease !important;
}
.stDownloadButton > button:hover {
    background: rgba(16,230,141,.14) !important;
    border-color: rgba(16,230,141,.45) !important;
}

/* ── dividers ── */
hr { border-color: rgba(255,255,255,0.06) !important; margin: 16px 0 !important; }

/* ── alerts ── */
[data-testid="stAlert"] {
    background: rgba(244,62,1,.08) !important;
    border: 1px solid rgba(244,62,1,.2) !important;
    border-radius: 8px !important;
    color: #f8f8f7 !important;
}
[data-testid="stAlert"][data-baseweb="notification"][kind="success"],
[data-testid="stNotification"] {
    background: rgba(16,230,141,.08) !important;
    border-color: rgba(16,230,141,.2) !important;
}

/* ── scrollbar ── */
::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: rgba(255,255,255,.1); border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def fmt_date(d):
    if not d:
        return ""
    if isinstance(d, (date, datetime)):
        return d.strftime("%d/%m/%Y")
    return str(d)

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

def sec(num_or_label, title, desc=""):
    """Section header — Groq style."""
    st.markdown(f"""
    <div style="margin:4px 0 28px">
      <div style="font-size:.6rem;font-weight:600;letter-spacing:.22em;text-transform:uppercase;
                  color:#766f6b;margin-bottom:10px">{num_or_label}</div>
      <div style="font-family:'Space Grotesk',sans-serif;font-size:1.35rem;font-weight:600;
                  color:#f8f8f7;letter-spacing:-.015em;line-height:1.2">{title}</div>
      {"" if not desc else f'<div style="font-size:.8rem;color:#766f6b;margin-top:6px;line-height:1.6">{desc}</div>'}
      <div style="width:100%;height:1px;background:rgba(255,255,255,.06);margin-top:20px"></div>
    </div>
    """, unsafe_allow_html=True)

def sub(label):
    st.markdown(f"""
    <div style="font-size:.6rem;font-weight:600;letter-spacing:.18em;text-transform:uppercase;
                color:#5a5450;margin:22px 0 10px;padding-left:1px">{label}</div>
    """, unsafe_allow_html=True)

def col_head(*labels, cols):
    for col, lbl in zip(cols, labels):
        col.markdown(
            f"<div style='font-size:.6rem;font-weight:600;letter-spacing:.14em;"
            f"text-transform:uppercase;color:#5a5450;padding-bottom:6px'>{lbl}</div>",
            unsafe_allow_html=True,
        )

def sim_nao(label, key):
    return st.selectbox(label, ["", "Sim", "Não"], key=key)

def nivel(label, key):
    return st.selectbox(label, ["", "1 — Baixo", "2 — Médio", "3 — Alto"], key=key)

def s_doc(label, key):
    return st.selectbox(label, ["", "Pendente", "Entregue", "Não se aplica"], key=key)

def s_lic(label, key):
    return st.selectbox(label, ["", "Válida", "Vencida", "Pendente", "Não se aplica"], key=key)

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style="padding:32px 0 24px">
      <div style="font-family:'Space Grotesk',sans-serif;font-size:1.5rem;font-weight:700;
                  letter-spacing:.12em;color:#f8f8f7">PRISMA</div>
      <div style="font-size:.58rem;color:#5a5450;letter-spacing:.14em;
                  text-transform:uppercase;margin-top:5px">Ficha Técnica do Cliente</div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    nav = [
        ("01", "Cadastro"),
        ("02", "Endereços"),
        ("03", "Contatos"),
        ("04", "Perfil"),
        ("05", "Comercial"),
        ("06", "Regras"),
        ("07", "Licenças"),
        ("08", "Estratégico"),
        ("09", "Crédito"),
    ]
    for num, label in nav:
        st.markdown(f"""
        <div style="display:flex;align-items:center;gap:12px;padding:6px 0">
          <span style="font-family:'IBM Plex Mono',monospace;font-size:.58rem;
                       color:#3a3430;font-weight:500">{num}</span>
          <span style="font-size:.78rem;color:#766f6b;font-weight:400">{label}</span>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    if st.button("Gerar Ficha", type="primary", use_container_width=True):
        st.session_state["_gerar"] = True

    if "arquivo_gerado" in st.session_state:
        st.download_button(
            "↓  Baixar Excel",
            data=st.session_state["arquivo_gerado"],
            file_name=st.session_state.get("_nome_arquivo", "PRISMA.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("""
        <div style="font-size:.65rem;color:#766f6b;text-align:center;margin-top:8px;
                    letter-spacing:.04em">Arquivo pronto para download</div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div style="position:fixed;bottom:28px;font-size:.6rem;color:#3a3430;
                letter-spacing:.06em">BIOTROP · uso interno</div>
    """, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tabs = st.tabs(["Cadastro", "Endereços", "Contatos", "Perfil",
                "Comercial", "Regras", "Licenças", "Estratégico", "Crédito"])

# ─────────────────────────────────────────────
# 01 · CADASTRO
# ─────────────────────────────────────────────
with tabs[0]:
    sec("Bloco 01", "Dados Cadastrais",
        "Identificação legal e fiscal do cliente.")

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

    c4, c5, c6 = st.columns([2, 2, 2])
    with c4:
        st.date_input("Data de Abertura", value=None, key="data_abertura")
    with c5:
        st.text_input("Grupo Econômico", key="grupo_economico")
    with c6:
        st.date_input("Cliente Desde", value=None, key="data_inicio_relacionamento")

    sub("Responsável Biotrop")

    rb1, rb2, rb3 = st.columns(3)
    with rb1: st.text_input("Nome completo", key="ctrl_responsavel_biotrop")
    with rb2: st.text_input("E-mail", key="ctrl_responsavel_biotrop_email")
    with rb3: st.text_input("Telefone", key="ctrl_responsavel_biotrop_tel")

# ─────────────────────────────────────────────
# 02 · ENDEREÇOS
# ─────────────────────────────────────────────
with tabs[1]:
    sec("Bloco 02", "Endereços",
        "Endereço fiscal e pontos de entrega.")

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
    entregas = {}
    h1, h2, h3, h4 = st.columns([1, 2, 3, 4])
    col_head("Opção", "Identificação", "Município / Estado", "Observações",
             cols=[h1, h2, h3, h4])

    for n in range(1, 5):
        ec1, ec2, ec3, ec4 = st.columns([1, 2, 3, 4])
        with ec1:
            st.markdown(f"<div style='padding-top:10px;font-family:IBM Plex Mono,monospace;"
                        f"font-size:.75rem;color:#5a5450'>0{n}</div>", unsafe_allow_html=True)
        with ec2:
            entregas[f"end_entrega_{n}_id"] = st.text_input(
                "", key=f"end_entrega_{n}_id", label_visibility="collapsed",
                placeholder="Nome da unidade")
        with ec3:
            entregas[f"end_entrega_{n}_municipio_estado"] = st.text_input(
                "", key=f"end_entrega_{n}_municipio_estado", label_visibility="collapsed",
                placeholder="Cidade / UF")
        with ec4:
            entregas[f"end_entrega_{n}_obs"] = st.text_input(
                "", key=f"end_entrega_{n}_obs", label_visibility="collapsed",
                placeholder="—")

# ─────────────────────────────────────────────
# 03 · CONTATOS
# ─────────────────────────────────────────────
with tabs[2]:
    sec("Bloco 03", "Contatos",
        "Dois contatos por área: comercial, financeiro, técnico e logística.")

    contatos = {}
    grupos = [
        ("Comercial",             "com"),
        ("Financeiro",            "fin"),
        ("Responsável Técnico",   "tec"),
        ("Logística",             "log"),
    ]
    for g_label, prefix in grupos:
        sub(g_label)
        h1, h2, h3, h4, h5 = st.columns([1, 3, 2, 2, 3])
        col_head("", "Nome", "Cargo", "Tel / Whatsapp", "E-mail",
                 cols=[h1, h2, h3, h4, h5])
        for idx in (1, 2):
            ic1, ic2, ic3, ic4, ic5 = st.columns([1, 3, 2, 2, 3])
            with ic1:
                st.markdown(f"<div style='padding-top:10px;font-family:IBM Plex Mono,monospace;"
                            f"font-size:.7rem;color:#5a5450'>{idx}°</div>",
                            unsafe_allow_html=True)
            with ic2:
                contatos[f"{prefix}{idx}_nome"]  = st.text_input("", key=f"{prefix}{idx}_nome",  label_visibility="collapsed", placeholder="—")
            with ic3:
                contatos[f"{prefix}{idx}_cargo"] = st.text_input("", key=f"{prefix}{idx}_cargo", label_visibility="collapsed", placeholder="—")
            with ic4:
                contatos[f"{prefix}{idx}_tel"]   = st.text_input("", key=f"{prefix}{idx}_tel",   label_visibility="collapsed", placeholder="—")
            with ic5:
                contatos[f"{prefix}{idx}_email"] = st.text_input("", key=f"{prefix}{idx}_email", label_visibility="collapsed", placeholder="—")

# ─────────────────────────────────────────────
# 04 · PERFIL
# ─────────────────────────────────────────────
with tabs[3]:
    sec("Bloco 04", "Perfil do Cliente",
        "Caracterização do negócio e relação com os produtos Biotrop.")

    p1, p2, p3 = st.columns([2, 2, 3])
    with p1:
        st.selectbox("Tipo de cliente", [
            "", "Distribuidor", "Revendedor", "Cooperativa",
            "Produtor Rural", "Importador", "Outro",
        ], key="tipo_cliente")
        st.text_input("Área Total (ha)", key="area_total_ha")
    with p2:
        st.text_input("Culturas Principais", key="culturas_principais")
        st.text_input("Região de Atuação", key="regiao_atuacao")
    with p3:
        st.text_area("Produtos Biotrop que já compra", height=112, key="produtos_utilizados",
                     placeholder="Liste os produtos já utilizados pelo cliente…")

# ─────────────────────────────────────────────
# 05 · COMERCIAL
# ─────────────────────────────────────────────
with tabs[4]:
    sec("Bloco 05", "Condições Comerciais e Financeiras",
        "Pagamento, crédito e política de bonificação.")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.text_input("Condição de Pagamento", placeholder="ex: 30/60/90 DDL", key="condicao_pagamento")
        st.text_input("Prazo Médio (dias)", key="prazo_medio_dias")
    with c2:
        st.text_input("Limite de Crédito (R$)", key="limite_credito")
        st.selectbox("Forma de Pagamento", [
            "", "Boleto", "PIX", "Transferência Bancária", "Cheque", "Cartão", "Misto",
        ], key="forma_pagamento")
    with c3:
        st.text_area("Política de Bonificação", height=112, key="politica_bonificacao")

    st.text_area("Condições Comerciais Especiais", height=72,
                 key="condicoes_especiais", placeholder="—")

# ─────────────────────────────────────────────
# 06 · REGRAS
# ─────────────────────────────────────────────
with tabs[5]:
    sec("Bloco 06", "Regras de Faturamento e Logística",
        "Requisitos operacionais para emissão de nota e entrega.")

    sub("Faturamento")
    f1, f2, f3, f4 = st.columns([2, 1, 1, 1])
    with f1: st.text_input("Data limite para faturamento", placeholder="ex: dia 20", key="fat_data_limite")
    with f2: sim_nao("Exige PO?",             "fat_exige_po")
    with f3: sim_nao("Exige Contrato?",        "fat_exige_contrato")
    with f4: sim_nao("Conferência de NF?",     "fat_conferencia_nf")

    f5, f6, f7 = st.columns([2, 2, 3])
    with f5: st.text_input("Prazo de Envio para NF", placeholder="ex: 48h antes", key="fat_prazo_nf")
    with f6: st.text_input("Shelf life mínimo", placeholder="ex: 12 meses", key="fat_shelf_life")
    with f7: st.text_area("Observações de Faturamento", height=72, key="fat_observacoes", placeholder="—")

    sub("Logística")
    l1, l2, l3, l4, l5 = st.columns([2, 1, 2, 1, 2])
    with l1: st.selectbox("Tipo de entrega", ["", "CIF", "FOB", "CIF e FOB"], key="log_tipo_entrega")
    with l2: sim_nao("Agendamento?",           "log_agendamento")
    with l3: st.text_input("Prazo mínimo agendamento", placeholder="ex: 48h", key="log_prazo_agendamento")
    with l4: sim_nao("Aviso?",                 "log_aviso_entrega")
    with l5: st.text_input("Antecedência aviso (h)", key="log_antecedencia_aviso")

    l6, l7, l8, l9 = st.columns([2, 2, 1, 1])
    with l6: st.text_input("Dias de recebimento", placeholder="ex: Seg a Sex", key="log_dias_recebimento")
    with l7: st.text_input("Horário de recebimento", placeholder="ex: 07h–17h", key="log_horario_recebimento")
    with l8: sim_nao("NF antecipada?",         "log_nf_antecipada")
    with l9: st.text_input("Antecedência NF (h)", key="log_antecedencia_nf")

    l10, l11, l12 = st.columns([1, 2, 4])
    with l10: sim_nao("Romaneio?",             "log_romaneio")
    with l11: st.text_input("Restrição transportadora", key="log_restricao_transportadora")
    with l12: st.text_area("Regras de acesso / EPI / Observações", height=72,
                            key="log_acesso_obs", placeholder="—")

# ─────────────────────────────────────────────
# 07 · LICENÇAS
# ─────────────────────────────────────────────
with tabs[6]:
    sec("Bloco 07", "Licenças e Documentação",
        "Status de validade das licenças regulatórias do cliente.")

    lic_data = {}
    licencas = [
        ("Licença Ambiental",   "lic_ambiental"),
        ("Licença de Operação", "lic_operacao"),
        ("Registro MAPA",       "lic_mapa"),
        ("Alvará Municipal",    "lic_alvara"),
        ("Certificação ISO",    "lic_iso"),
    ]

    hd = st.columns([3, 2, 2, 4])
    col_head("Documento", "Status", "Validade", "Observações", cols=hd)

    for nome_lic, prefix in licencas:
        c1, c2, c3, c4 = st.columns([3, 2, 2, 4])
        with c1:
            st.markdown(
                f"<div style='padding-top:10px;font-size:.85rem;color:#a09890'>{nome_lic}</div>",
                unsafe_allow_html=True)
        with c2:
            lic_data[f"{prefix}_status"]   = s_lic("", f"{prefix}_status")
        with c3:
            lic_data[f"{prefix}_validade"] = st.date_input(
                "", value=None, key=f"{prefix}_validade", label_visibility="collapsed")
        with c4:
            lic_data[f"{prefix}_obs"]      = st.text_input(
                "", key=f"{prefix}_obs", label_visibility="collapsed", placeholder="—")

# ─────────────────────────────────────────────
# 08 · ESTRATÉGICO
# ─────────────────────────────────────────────
with tabs[7]:
    sec("Bloco 08", "Informações Estratégicas e Complexidade",
        "Inteligência competitiva e avaliação operacional do cliente.")

    sub("Informações Estratégicas")
    es1, es2, es3 = st.columns(3)
    with es1:
        st.text_area("Concorrentes utilizados", height=100, key="estrat_concorrentes", placeholder="—")
        st.text_input("Potencial de volume estimado", placeholder="ex: 500 t/ano", key="estrat_potencial_volume")
    with es2:
        st.text_area("Histórico de relacionamentos", height=100, key="estrat_historico", placeholder="—")
        st.text_input("Participação estimada (%)", placeholder="ex: 30%", key="estrat_participacao")
    with es3:
        st.selectbox("Classificação de risco", ["", "Baixo", "Médio", "Alto"], key="estrat_risco")
        st.text_area("Observações estratégicas", height=100, key="estrat_observacoes", placeholder="—")

    sub("Complexidade Operacional")
    co1, co2 = st.columns([2, 5])
    with co1:
        st.selectbox("Classificação Geral", ["", "Baixa", "Média", "Alta"], key="compl_classificacao")
    with co2:
        st.text_area("Justificativa", height=72, key="compl_justificativa", placeholder="—")

    hd = st.columns([4, 1, 4])
    col_head("Critério", "Nível (1–3)", "Observação", cols=hd)

    compl_data = {}
    criterios = [
        ("Exigências logísticas",          "compl_log"),
        ("Burocracia de faturamento",       "compl_fat"),
        ("Dificuldade de acesso",           "compl_acesso"),
        ("Nível de exigência operacional",  "compl_operacional"),
    ]
    for nome_crit, prefix in criterios:
        cr1, cr2, cr3 = st.columns([4, 1, 4])
        with cr1:
            st.markdown(
                f"<div style='padding-top:10px;font-size:.85rem;color:#a09890'>{nome_crit}</div>",
                unsafe_allow_html=True)
        with cr2:
            compl_data[f"{prefix}_nivel"] = nivel("", f"{prefix}_nivel")
        with cr3:
            compl_data[f"{prefix}_obs"]   = st.text_input(
                "", key=f"{prefix}_obs", label_visibility="collapsed", placeholder="—")

# ─────────────────────────────────────────────
# 09 · CRÉDITO + CONTROLE
# ─────────────────────────────────────────────
with tabs[8]:
    cred_data = {}

    def render_docs(prefix, bloco, titulo, docs):
        sec(f"Bloco {bloco}", titulo)
        hd = st.columns([5, 2, 2, 3])
        col_head("Documento", "Status", "Data de entrega", "Observações", cols=hd)
        for n, nome_doc in enumerate(docs, start=1):
            c1, c2, c3, c4 = st.columns([5, 2, 2, 3])
            with c1:
                st.markdown(
                    f"<div style='padding-top:10px;font-size:.82rem;color:#a09890'>{nome_doc}</div>",
                    unsafe_allow_html=True)
            with c2:
                cred_data[f"{prefix}_doc{n}_status"] = s_doc("", f"{prefix}_doc{n}_status")
            with c3:
                cred_data[f"{prefix}_doc{n}_data"]   = st.date_input(
                    "", value=None, key=f"{prefix}_doc{n}_data", label_visibility="collapsed")
            with c4:
                cred_data[f"{prefix}_doc{n}_obs"]    = st.text_input(
                    "", key=f"{prefix}_doc{n}_obs", label_visibility="collapsed", placeholder="—")

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
        st.selectbox("Status do cadastro", [
            "", "Em preenchimento", "Completo", "Em revisão", "Aprovado",
        ], key="ctrl_status")

# ══════════════════════════════════════════════════════════════════════════════
# GERAR FICHA
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.get("_gerar"):
    st.session_state["_gerar"] = False
    razao = st.session_state.get("razao_social", "").strip()

    if not razao:
        st.sidebar.error("Preencha a Razão Social antes de gerar.")
    else:
        log_obs = st.session_state.get("log_acesso_obs", "")
        data = {
            "razao_social":               razao,
            "nome_fantasia":              st.session_state.get("nome_fantasia", ""),
            "cnpj":                       st.session_state.get("cnpj", ""),
            "inscricao_estadual":         st.session_state.get("inscricao_estadual", ""),
            "inscricao_municipal":        st.session_state.get("inscricao_municipal", ""),
            "cnae":                       st.session_state.get("cnae", ""),
            "data_abertura":              fmt_date(st.session_state.get("data_abertura")),
            "grupo_economico":            st.session_state.get("grupo_economico", ""),
            "data_inicio_relacionamento": fmt_date(st.session_state.get("data_inicio_relacionamento")),
            "ctrl_responsavel_biotrop":          st.session_state.get("ctrl_responsavel_biotrop", ""),
            "ctrl_responsavel_biotrop-email":    st.session_state.get("ctrl_responsavel_biotrop_email", ""),
            "ctrl_responsavel_biotrop-telefone": st.session_state.get("ctrl_responsavel_biotrop_tel", ""),
            "end_fiscal_logradouro":  st.session_state.get("end_fiscal_logradouro", ""),
            "end_fiscal_numero":      st.session_state.get("end_fiscal_numero", ""),
            "end_fiscal_complemento": st.session_state.get("end_fiscal_complemento", ""),
            "end_fiscal_bairro":      st.session_state.get("end_fiscal_bairro", ""),
            "end_fiscal_municipio":   st.session_state.get("end_fiscal_municipio", ""),
            "end_fiscal_estado":      st.session_state.get("end_fiscal_estado", ""),
            "end_fiscal_cep":         st.session_state.get("end_fiscal_cep", ""),
            **{k: st.session_state.get(k, "") for k in contatos},
            "tipo_cliente":        st.session_state.get("tipo_cliente", ""),
            "culturas_principais": st.session_state.get("culturas_principais", ""),
            "area_total_ha":       st.session_state.get("area_total_ha", ""),
            "regiao_atuacao":      st.session_state.get("regiao_atuacao", ""),
            "produtos_utilizados": st.session_state.get("produtos_utilizados", ""),
            "condicao_pagamento":   st.session_state.get("condicao_pagamento", ""),
            "prazo_medio_dias":     st.session_state.get("prazo_medio_dias", ""),
            "limite_credito":       st.session_state.get("limite_credito", ""),
            "forma_pagamento":      st.session_state.get("forma_pagamento", ""),
            "politica_bonificacao": st.session_state.get("politica_bonificacao", ""),
            "condicoes_especiais":  st.session_state.get("condicoes_especiais", ""),
            "fat_data_limite":    st.session_state.get("fat_data_limite", ""),
            "fat_exige_po":       st.session_state.get("fat_exige_po", ""),
            "fat_exige_contrato": st.session_state.get("fat_exige_contrato", ""),
            "fat_conferencia_nf": st.session_state.get("fat_conferencia_nf", ""),
            "fat_prazo_nf":       st.session_state.get("fat_prazo_nf", ""),
            "fat_shelf_life":     st.session_state.get("fat_shelf_life", ""),
            "fat_observacoes":    st.session_state.get("fat_observacoes", ""),
            "log_tipo_entrega":             st.session_state.get("log_tipo_entrega", ""),
            "log_agendamento":              st.session_state.get("log_agendamento", ""),
            "log_prazo_agendamento":        st.session_state.get("log_prazo_agendamento", ""),
            "log_dias_recebimento":         st.session_state.get("log_dias_recebimento", ""),
            "log_horario_recebimento":      st.session_state.get("log_horario_recebimento", ""),
            "log_aviso_entrega":            st.session_state.get("log_aviso_entrega", ""),
            "log_antecedencia_aviso":       st.session_state.get("log_antecedencia_aviso", ""),
            "log_nf_antecipada":            st.session_state.get("log_nf_antecipada", ""),
            "log_antecedencia_nf":          st.session_state.get("log_antecedencia_nf", ""),
            "log_romaneio":                 st.session_state.get("log_romaneio", ""),
            "log_restricao_transportadora": st.session_state.get("log_restricao_transportadora", ""),
            "log_regras_acesso":            log_obs,
            "log_observacoes":              log_obs,
            **{k: fmt_date(v) if isinstance(v, (date, datetime)) else (v or "")
               for k, v in lic_data.items()},
            "estrat_concorrentes":     st.session_state.get("estrat_concorrentes", ""),
            "estrat_potencial_volume": st.session_state.get("estrat_potencial_volume", ""),
            "estrat_participacao":     st.session_state.get("estrat_participacao", ""),
            "estrat_historico":        st.session_state.get("estrat_historico", ""),
            "estrat_risco":            st.session_state.get("estrat_risco", ""),
            "estrat_observacoes":      st.session_state.get("estrat_observacoes", ""),
            "compl_classificacao": st.session_state.get("compl_classificacao", ""),
            "compl_justificativa": st.session_state.get("compl_justificativa", ""),
            **{k: (v or "") for k, v in compl_data.items()},
            **{k: fmt_date(v) if isinstance(v, (date, datetime)) else (v or "")
               for k, v in cred_data.items()},
            "ctrl_cadastrado_por":          st.session_state.get("ctrl_cadastrado_por", ""),
            "ctrl_data_cadastro":           fmt_date(st.session_state.get("ctrl_data_cadastro")),
            "ctrl_ultima_atualizacao":      fmt_date(st.session_state.get("ctrl_ultima_atualizacao")),
            "ctrl_responsavel_atualizacao": st.session_state.get("ctrl_responsavel_atualizacao", ""),
            "ctrl_status":                  st.session_state.get("ctrl_status", ""),
        }
        for n in range(1, 5):
            for campo in ("id", "municipio_estado", "obs"):
                data[f"end_entrega_{n}_{campo}"] = st.session_state.get(
                    f"end_entrega_{n}_{campo}", "")

        try:
            st.session_state["arquivo_gerado"]  = fill_template(data)
            st.session_state["_nome_arquivo"]   = f"PRISMA_{razao.replace(' ', '_')}.xlsx"
            st.rerun()
        except FileNotFoundError:
            st.sidebar.error("Template não encontrado.")
        except Exception as e:
            st.sidebar.error(f"Erro: {e}")

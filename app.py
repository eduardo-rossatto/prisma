import streamlit as st
import openpyxl
from io import BytesIO
from datetime import date, datetime

st.set_page_config(page_title="PRISMA · Ficha Técnica do Cliente", layout="wide", page_icon="🤝")

TEMPLATE_PATH = "template/Prisma - Template.xlsx"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }

#MainMenu, footer, header { visibility: hidden; }

.stApp { background-color: #f5f5f5; }

/* sidebar */
[data-testid="stSidebar"] {
    background-color: #0a0a0a !important;
    padding-top: 0 !important;
}
[data-testid="stSidebar"] * { color: #ffffff !important; }
[data-testid="stSidebar"] hr { border-color: #2a2a2a !important; }
[data-testid="stSidebar"] .stButton > button {
    background-color: #ffffff !important;
    color: #0a0a0a !important;
    border: none !important;
    font-weight: 600 !important;
    letter-spacing: 0.06em !important;
    width: 100%;
}
[data-testid="stSidebar"] .stButton > button:hover {
    background-color: #e5e5e5 !important;
}
[data-testid="stSidebar"] .stDownloadButton > button {
    background-color: #1a7a4a !important;
    color: #ffffff !important;
    border: none !important;
    font-weight: 600 !important;
    width: 100%;
}

/* tabs */
.stTabs [data-baseweb="tab-list"] {
    gap: 2px;
    background-color: #ebebeb;
    border-radius: 10px;
    padding: 4px;
    border-bottom: none !important;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 7px;
    font-size: 0.78rem;
    font-weight: 500;
    letter-spacing: 0.03em;
    color: #666 !important;
    border: none !important;
    background: transparent !important;
    padding: 7px 16px !important;
}
.stTabs [aria-selected="true"] {
    background-color: #ffffff !important;
    color: #0a0a0a !important;
    box-shadow: 0 1px 4px rgba(0,0,0,0.1) !important;
}
.stTabs [data-baseweb="tab-highlight"] { display: none; }
.stTabs [data-baseweb="tab-border"]    { display: none; }

/* section card */
.sec-card {
    background: #ffffff;
    border-radius: 12px;
    border: 1px solid #e8e8e8;
    padding: 24px 28px 20px;
    margin-bottom: 20px;
}
.sec-label {
    font-size: 0.6rem;
    font-weight: 500;
    letter-spacing: 0.25em;
    text-transform: uppercase;
    color: #aaaaaa;
    margin-bottom: 4px;
}
.sec-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: #0a0a0a;
    margin-bottom: 16px;
    padding-bottom: 12px;
    border-bottom: 1px solid #f0f0f0;
}
.sub-label {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #999999;
    margin: 16px 0 8px;
}
.table-header {
    display: grid;
    gap: 8px;
    padding: 8px 4px;
    border-bottom: 1px solid #eeeeee;
    margin-bottom: 4px;
}
.table-header span {
    font-size: 0.68rem;
    font-weight: 600;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #aaaaaa;
}

/* inputs */
[data-testid="stTextInput"] label,
[data-testid="stSelectbox"] label,
[data-testid="stDateInput"] label,
[data-testid="stTextArea"] label,
[data-testid="stNumberInput"] label {
    font-size: 0.75rem !important;
    font-weight: 500 !important;
    color: #555555 !important;
    letter-spacing: 0.02em !important;
}
input, textarea, [data-baseweb="select"] > div {
    border-radius: 7px !important;
    border-color: #e0e0e0 !important;
    font-size: 0.85rem !important;
}
input:focus, textarea:focus { border-color: #0a0a0a !important; box-shadow: none !important; }
</style>
""", unsafe_allow_html=True)

# ── helpers ───────────────────────────────────────────────────────────────────

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

def card_open(label, title):
    st.markdown(f"""
    <div class="sec-card">
      <div class="sec-label">{label}</div>
      <div class="sec-title">{title}</div>
    </div>
    """, unsafe_allow_html=True)

def sub(text):
    st.markdown(f'<div class="sub-label">{text}</div>', unsafe_allow_html=True)

def sim_nao(label, key):
    return st.selectbox(label, ["", "Sim", "Não"], key=key)

def nivel(label, key):
    return st.selectbox(label, ["", "1 — Baixo", "2 — Médio", "3 — Alto"], key=key)

def status_doc(label, key):
    return st.selectbox(label, ["", "Pendente", "Entregue", "Não se aplica"], key=key)

def status_lic(label, key):
    return st.selectbox(label, ["", "Válida", "Vencida", "Pendente", "Não se aplica"], key=key)

# ── sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding:32px 0 8px;text-align:center'>
      <div style='font-size:1.8rem;font-weight:200;letter-spacing:.2em'>PRISMA</div>
      <div style='font-size:0.6rem;color:#555;letter-spacing:.1em;margin-top:4px'>
        FICHA TÉCNICA DO CLIENTE
      </div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    st.markdown("""
    <div style='font-size:0.65rem;color:#555;letter-spacing:.12em;text-transform:uppercase;margin-bottom:12px'>
      Seções
    </div>
    """, unsafe_allow_html=True)

    secoes = [
        "Cadastro",
        "Endereços",
        "Contatos",
        "Perfil",
        "Comercial",
        "Regras",
        "Licenças",
        "Estratégico",
        "Crédito",
    ]
    for s in secoes:
        st.markdown(f"<div style='font-size:0.8rem;padding:4px 0;color:#888'>· {s}</div>",
                    unsafe_allow_html=True)

    st.divider()

    if st.button("Gerar Ficha", use_container_width=True):
        st.session_state["_gerar"] = True

    if "arquivo_gerado" in st.session_state:
        nome = st.session_state.get("nome_arquivo", "PRISMA.xlsx")
        st.download_button(
            "⬇ Baixar Excel",
            data=st.session_state["arquivo_gerado"],
            file_name=nome,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    st.markdown("""
    <div style='font-size:0.62rem;color:#444;text-align:center;padding-bottom:16px'>
      BIOTROP · uso interno
    </div>
    """, unsafe_allow_html=True)

# ── tabs ──────────────────────────────────────────────────────────────────────
tabs = st.tabs(["Cadastro", "Endereços", "Contatos", "Perfil",
                "Comercial", "Regras", "Licenças", "Estratégico", "Crédito"])

# ═══════════════════════════════════════════════════════
# 1 · CADASTRO
# ═══════════════════════════════════════════════════════
with tabs[0]:
    st.markdown('<div class="sec-card"><div class="sec-label">Bloco 1</div><div class="sec-title">Dados Cadastrais</div></div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Razão Social", key="razao_social")
        st.text_input("CNPJ", placeholder="00.000.000/0000-00", key="cnpj")
        st.text_input("Inscrição Estadual", key="inscricao_estadual")
        st.text_input("Inscrição Municipal", key="inscricao_municipal")
    with c2:
        st.text_input("Nome Fantasia", key="nome_fantasia")
        st.text_input("CNAE", placeholder="0000-0/00", key="cnae")
        st.date_input("Data de Abertura", value=None, key="data_abertura")
        st.text_input("Grupo Econômico", key="grupo_economico")

    st.date_input("Cliente Desde", value=None, key="data_inicio_relacionamento",
                  help="Data de início do relacionamento com a Biotrop")

    st.markdown('<div class="sec-card" style="margin-top:20px"><div class="sec-label">Responsável</div><div class="sec-title">Responsável Biotrop pelo Cliente</div></div>', unsafe_allow_html=True)

    rb1, rb2, rb3 = st.columns(3)
    with rb1:
        st.text_input("Nome completo", key="ctrl_responsavel_biotrop")
    with rb2:
        st.text_input("E-mail", key="ctrl_responsavel_biotrop_email")
    with rb3:
        st.text_input("Telefone", key="ctrl_responsavel_biotrop_tel")

# ═══════════════════════════════════════════════════════
# 2 · ENDEREÇOS
# ═══════════════════════════════════════════════════════
with tabs[1]:
    st.markdown('<div class="sec-card"><div class="sec-label">Bloco 2</div><div class="sec-title">Endereços</div></div>', unsafe_allow_html=True)

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
    for n in range(1, 5):
        ec1, ec2, ec3 = st.columns([2, 3, 4])
        with ec1:
            entregas[f"end_entrega_{n}_id"] = st.text_input(
                f"Unidade {n}", key=f"end_entrega_{n}_id")
        with ec2:
            entregas[f"end_entrega_{n}_municipio_estado"] = st.text_input(
                "Município / Estado", key=f"end_entrega_{n}_municipio_estado",
                label_visibility="visible" if n == 1 else "collapsed")
        with ec3:
            entregas[f"end_entrega_{n}_obs"] = st.text_input(
                "Observações", key=f"end_entrega_{n}_obs",
                label_visibility="visible" if n == 1 else "collapsed")

# ═══════════════════════════════════════════════════════
# 3 · CONTATOS
# ═══════════════════════════════════════════════════════
with tabs[2]:
    st.markdown('<div class="sec-card"><div class="sec-label">Bloco 3</div><div class="sec-title">Contatos</div></div>', unsafe_allow_html=True)

    contatos = {}
    grupos = [
        ("Contato Comercial",           "com"),
        ("Contato Financeiro",          "fin"),
        ("Responsável Técnico",         "tec"),
        ("Responsável pela Logística",  "log"),
    ]

    for g_label, prefix in grupos:
        sub(g_label)
        for idx in (1, 2):
            cc1, cc2, cc3, cc4 = st.columns(4)
            lv = "visible" if idx == 1 else "collapsed"
            with cc1:
                contatos[f"{prefix}{idx}_nome"]  = st.text_input("Nome",          key=f"{prefix}{idx}_nome",  label_visibility=lv)
            with cc2:
                contatos[f"{prefix}{idx}_cargo"] = st.text_input("Cargo",         key=f"{prefix}{idx}_cargo", label_visibility=lv)
            with cc3:
                contatos[f"{prefix}{idx}_tel"]   = st.text_input("Tel/Whatsapp",  key=f"{prefix}{idx}_tel",   label_visibility=lv)
            with cc4:
                contatos[f"{prefix}{idx}_email"] = st.text_input("E-mail",        key=f"{prefix}{idx}_email", label_visibility=lv)

# ═══════════════════════════════════════════════════════
# 4 · PERFIL
# ═══════════════════════════════════════════════════════
with tabs[3]:
    st.markdown('<div class="sec-card"><div class="sec-label">Bloco 4</div><div class="sec-title">Perfil do Cliente</div></div>', unsafe_allow_html=True)

    p1, p2 = st.columns(2)
    with p1:
        st.selectbox("Tipo de cliente", [
            "", "Distribuidor", "Revendedor", "Cooperativa",
            "Produtor Rural", "Importador", "Outro",
        ], key="tipo_cliente")
        st.text_input("Área Total (ha)", key="area_total_ha")
        st.text_input("Região de Atuação", key="regiao_atuacao")
    with p2:
        st.text_input("Culturas Principais", key="culturas_principais")
        st.text_area("Produtos Biotrop que já compra", height=120, key="produtos_utilizados")

# ═══════════════════════════════════════════════════════
# 5 · COMERCIAL
# ═══════════════════════════════════════════════════════
with tabs[4]:
    st.markdown('<div class="sec-card"><div class="sec-label">Bloco 5</div><div class="sec-title">Condições Comerciais e Financeiras</div></div>', unsafe_allow_html=True)

    cc1, cc2 = st.columns(2)
    with cc1:
        st.text_input("Condição de Pagamento", placeholder="ex: 30/60/90 DDL", key="condicao_pagamento")
        st.text_input("Prazo Médio (dias)", key="prazo_medio_dias")
        st.text_input("Limite de Crédito (R$)", key="limite_credito")
    with cc2:
        st.selectbox("Forma de Pagamento", [
            "", "Boleto", "PIX", "Transferência Bancária", "Cheque", "Cartão", "Misto",
        ], key="forma_pagamento")
        st.text_area("Política de Bonificação", height=106, key="politica_bonificacao")

    st.text_area("Condições Comerciais Especiais", height=80, key="condicoes_especiais")

# ═══════════════════════════════════════════════════════
# 6 · REGRAS (Faturamento + Logística)
# ═══════════════════════════════════════════════════════
with tabs[5]:
    st.markdown('<div class="sec-card"><div class="sec-label">Bloco 6</div><div class="sec-title">Regras de Faturamento</div></div>', unsafe_allow_html=True)

    f1, f2 = st.columns(2)
    with f1:
        st.text_input("Data limite para faturamento", placeholder="ex: dia 20", key="fat_data_limite")
        sim_nao("Exige pedido formal (PO)?",         "fat_exige_po")
        sim_nao("Exige Contrato?",                   "fat_exige_contrato")
        sim_nao("Exige Conferência Prévia de NF?",   "fat_conferencia_nf")
    with f2:
        st.text_input("Prazo de Envio para NF", placeholder="ex: 48h antes do embarque", key="fat_prazo_nf")
        st.text_input("Shelf life mínimo do cliente", placeholder="ex: 12 meses", key="fat_shelf_life")
        st.text_area("Observações de Faturamento", height=94, key="fat_observacoes")

    st.markdown('<div class="sec-card" style="margin-top:20px"><div class="sec-label">Bloco 7</div><div class="sec-title">Regras Logísticas</div></div>', unsafe_allow_html=True)

    l1, l2 = st.columns(2)
    with l1:
        st.selectbox("Tipo de entrega", ["", "CIF", "FOB", "CIF e FOB"], key="log_tipo_entrega")
        sim_nao("Necessita Agendamento?",             "log_agendamento")
        st.text_input("Prazo Mínimo para Agendamento", placeholder="ex: 48h", key="log_prazo_agendamento")
        st.text_input("Dias de Recebimento",           placeholder="ex: Seg a Sex", key="log_dias_recebimento")
        st.text_input("Horário de Recebimento",        placeholder="ex: 07h às 17h", key="log_horario_recebimento")
        sim_nao("Avisar antes da Entrega?",            "log_aviso_entrega")
        st.text_input("Antecedência para Aviso (h)",   key="log_antecedencia_aviso")
    with l2:
        sim_nao("Exige Envio Antecipado de NF?",       "log_nf_antecipada")
        st.text_input("Antecedência de NF (h)",        key="log_antecedencia_nf")
        sim_nao("Exige Romaneio?",                     "log_romaneio")
        st.text_input("Restrição de Transportadora",   key="log_restricao_transportadora")
        st.text_area("Regras de Acesso (EPI, docs)",   height=80,  key="log_regras_acesso")
        st.text_area("Observações Logísticas",         height=80,  key="log_observacoes")

# ═══════════════════════════════════════════════════════
# 7 · LICENÇAS
# ═══════════════════════════════════════════════════════
with tabs[6]:
    st.markdown('<div class="sec-card"><div class="sec-label">Bloco 8</div><div class="sec-title">Licenças e Documentação</div></div>', unsafe_allow_html=True)

    lic_data = {}
    licencas = [
        ("Licença Ambiental",   "lic_ambiental"),
        ("Licença de Operação", "lic_operacao"),
        ("Registro MAPA",       "lic_mapa"),
        ("Alvará Municipal",    "lic_alvara"),
        ("Certificação ISO",    "lic_iso"),
    ]

    hd1, hd2, hd3, hd4 = st.columns([3, 2, 2, 3])
    hd1.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Documento</div>", unsafe_allow_html=True)
    hd2.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Status</div>", unsafe_allow_html=True)
    hd3.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Validade</div>", unsafe_allow_html=True)
    hd4.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Observações</div>", unsafe_allow_html=True)

    for nome_lic, prefix in licencas:
        c1, c2, c3, c4 = st.columns([3, 2, 2, 3])
        with c1:
            st.markdown(f"<div style='padding-top:10px;font-size:.85rem;color:#333'>{nome_lic}</div>", unsafe_allow_html=True)
        with c2:
            lic_data[f"{prefix}_status"]   = status_lic("", f"{prefix}_status")
        with c3:
            lic_data[f"{prefix}_validade"] = st.date_input("", value=None, key=f"{prefix}_validade", label_visibility="collapsed")
        with c4:
            lic_data[f"{prefix}_obs"]      = st.text_input("", key=f"{prefix}_obs", label_visibility="collapsed")

# ═══════════════════════════════════════════════════════
# 8 · ESTRATÉGICO (Informações + Complexidade)
# ═══════════════════════════════════════════════════════
with tabs[7]:
    st.markdown('<div class="sec-card"><div class="sec-label">Bloco 9</div><div class="sec-title">Informações Estratégicas</div></div>', unsafe_allow_html=True)

    es1, es2 = st.columns(2)
    with es1:
        st.text_area("Concorrentes que o cliente utiliza", height=100, key="estrat_concorrentes")
        st.text_input("Potencial de Volume Estimado", placeholder="ex: 500 t/ano", key="estrat_potencial_volume")
        st.text_input("Participação Estimada (%)", placeholder="ex: 30%", key="estrat_participacao")
    with es2:
        st.text_area("Histórico de Relacionamentos", height=100, key="estrat_historico")
        st.selectbox("Classificação de Risco", ["", "Baixo", "Médio", "Alto"], key="estrat_risco")
        st.text_area("Observações Estratégicas", height=80, key="estrat_observacoes")

    st.markdown('<div class="sec-card" style="margin-top:20px"><div class="sec-label">Bloco 10</div><div class="sec-title">Complexidade Operacional</div></div>', unsafe_allow_html=True)

    co1, co2 = st.columns(2)
    with co1:
        st.selectbox("Classificação Geral", ["", "Baixa", "Média", "Alta"], key="compl_classificacao")
    with co2:
        st.text_area("Justificativa", height=80, key="compl_justificativa")

    sub("Critérios de Avaliação — 1 = Baixo · 2 = Médio · 3 = Alto")

    hd1, hd2, hd3 = st.columns([4, 1, 4])
    hd1.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Critério</div>", unsafe_allow_html=True)
    hd2.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Nível</div>", unsafe_allow_html=True)
    hd3.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Observação</div>", unsafe_allow_html=True)

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
            st.markdown(f"<div style='padding-top:10px;font-size:.85rem;color:#333'>{nome_crit}</div>", unsafe_allow_html=True)
        with cr2:
            compl_data[f"{prefix}_nivel"] = nivel("", f"{prefix}_nivel")
        with cr3:
            compl_data[f"{prefix}_obs"]   = st.text_input("", key=f"{prefix}_obs", label_visibility="collapsed", placeholder="Observação")

# ═══════════════════════════════════════════════════════
# 9 · CRÉDITO + CONTROLE
# ═══════════════════════════════════════════════════════
with tabs[8]:
    cred_data = {}

    def render_docs(prefix, titulo, bloco, docs):
        st.markdown(f'<div class="sec-card"><div class="sec-label">Bloco {bloco}</div><div class="sec-title">{titulo}</div></div>', unsafe_allow_html=True)
        hd1, hd2, hd3, hd4 = st.columns([5, 2, 2, 3])
        hd1.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Documento</div>", unsafe_allow_html=True)
        hd2.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Status</div>", unsafe_allow_html=True)
        hd3.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Data de Entrega</div>", unsafe_allow_html=True)
        hd4.markdown("<div style='font-size:.68rem;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:#aaa'>Observações</div>", unsafe_allow_html=True)
        for n, nome_doc in enumerate(docs, start=1):
            c1, c2, c3, c4 = st.columns([5, 2, 2, 3])
            with c1:
                st.markdown(f"<div style='padding-top:10px;font-size:.82rem;color:#333'>{nome_doc}</div>", unsafe_allow_html=True)
            with c2:
                cred_data[f"{prefix}_doc{n}_status"] = status_doc("", f"{prefix}_doc{n}_status")
            with c3:
                cred_data[f"{prefix}_doc{n}_data"] = st.date_input("", value=None, key=f"{prefix}_doc{n}_data", label_visibility="collapsed")
            with c4:
                cred_data[f"{prefix}_doc{n}_obs"] = st.text_input("", key=f"{prefix}_doc{n}_obs", label_visibility="collapsed")

    render_docs("cred_ltda", "Análise de Crédito — Empresas LTDA", "11", [
        "Documentos pessoais dos sócios e cônjuges (RG / CPF)",
        "Comprovante de endereço (sócios)",
        "Certidão de casamento dos sócios (se aplicável)",
        "Contrato social e última alteração contratual",
        "Balanço Patrimonial e DRE — últimos 2 anos (assinados)",
        "Imposto de Renda dos sócios — último exercício",
    ])

    st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)

    render_docs("cred_coop", "Análise de Crédito — Cooperativas / S.A. / Usinas", "12", [
        "Documentos pessoais dos dirigentes",
        "Comprovante de endereço dos dirigentes",
        "Certidão de casamento dos sócios (se aplicável)",
        "Estatuto Social",
        "Ata de Eleição da Diretoria vigente",
        "Balanço Patrimonial e DRE — últimos 2 anos (assinados)",
    ])

    st.markdown('<div class="sec-card" style="margin-top:20px"><div class="sec-label">Controle</div><div class="sec-title">Controle Interno</div></div>', unsafe_allow_html=True)

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

# ═══════════════════════════════════════════════════════
# GERAR FICHA (disparado pelo sidebar)
# ═══════════════════════════════════════════════════════
if st.session_state.get("_gerar"):
    st.session_state["_gerar"] = False

    razao = st.session_state.get("razao_social", "").strip()
    if not razao:
        st.sidebar.error("Preencha a Razão Social antes de gerar.")
    else:
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
            "log_regras_acesso":            st.session_state.get("log_regras_acesso", ""),
            "log_observacoes":              st.session_state.get("log_observacoes", ""),
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
                k = f"end_entrega_{n}_{campo}"
                data[k] = st.session_state.get(k, "")

        try:
            st.session_state["arquivo_gerado"] = fill_template(data)
            st.session_state["nome_arquivo"] = f"PRISMA_{razao.replace(' ', '_')}.xlsx"
            st.sidebar.success("Pronto! Clique em Baixar Excel.")
            st.rerun()
        except FileNotFoundError:
            st.sidebar.error("Template não encontrado.")
        except Exception as e:
            st.sidebar.error(f"Erro: {e}")

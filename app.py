import streamlit as st
import openpyxl
from io import BytesIO
from datetime import date, datetime

st.set_page_config(page_title="PRISMA", layout="wide", page_icon="🌿")

TEMPLATE_PATH = "template/Prisma - Template.xlsx"

# ── helpers ──────────────────────────────────────────────────────────────────

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

def sim_nao(label, key):
    return st.selectbox(label, ["", "Sim", "Não"], key=key)

def nivel(label, key):
    return st.selectbox(label, ["", "1 — Baixo", "2 — Médio", "3 — Alto"], key=key)

def status_doc(label, key):
    return st.selectbox(label, ["", "Pendente", "Entregue", "Não se aplica"], key=key)

def status_lic(label, key):
    return st.selectbox(label, ["", "Válida", "Vencida", "Pendente", "Não se aplica"], key=key)

# ── cabeçalho ─────────────────────────────────────────────────────────────────

st.markdown("""
<div style='text-align:center;padding:24px 0 8px'>
  <h1 style='font-size:2.4rem;letter-spacing:.15em;margin:0'>PRISMA</h1>
  <p style='color:#6b7280;font-size:.85rem;letter-spacing:.08em;margin:4px 0 0'>
    Partner · Relationship · Intelligence · Strategy · Mapping · Assessment
  </p>
</div>
""", unsafe_allow_html=True)

st.divider()

# ── tabs ──────────────────────────────────────────────────────────────────────

tabs = st.tabs([
    "1 · Cadastro",
    "2 · Endereços",
    "3 · Contatos",
    "4 · Perfil & Comercial",
    "5 · Faturamento & Logística",
    "6 · Licenças",
    "7 · Estratégico",
    "8 · Crédito",
    "9 · Controle",
])

# ═══════════════════════════════════════════════════════════════════
# TAB 1 — CADASTRO
# ═══════════════════════════════════════════════════════════════════
with tabs[0]:
    st.subheader("1. Dados Cadastrais")
    c1, c2 = st.columns(2)
    with c1:
        razao_social           = st.text_input("Razão Social *", key="razao_social")
        cnpj                   = st.text_input("CNPJ", placeholder="00.000.000/0000-00", key="cnpj")
        inscricao_estadual     = st.text_input("Inscrição Estadual", key="inscricao_estadual")
        inscricao_municipal    = st.text_input("Inscrição Municipal", key="inscricao_municipal")
    with c2:
        nome_fantasia          = st.text_input("Nome Fantasia *", key="nome_fantasia")
        cnae                   = st.text_input("CNAE", placeholder="0000-0/00", key="cnae")
        data_abertura          = st.date_input("Data de Abertura", value=None, key="data_abertura")
        grupo_economico        = st.text_input("Grupo Econômico", key="grupo_economico")

    data_inicio_relacionamento = st.date_input("Cliente Desde", value=None, key="data_inicio_relacionamento")

    st.divider()
    st.subheader("Responsável Biotrop pelo Cliente")
    rb1, rb2, rb3 = st.columns(3)
    with rb1:
        ctrl_responsavel_biotrop         = st.text_input("Nome completo *", key="ctrl_responsavel_biotrop")
    with rb2:
        ctrl_responsavel_biotrop_email   = st.text_input("E-mail", key="ctrl_responsavel_biotrop_email")
    with rb3:
        ctrl_responsavel_biotrop_tel     = st.text_input("Telefone", key="ctrl_responsavel_biotrop_tel")

# ═══════════════════════════════════════════════════════════════════
# TAB 2 — ENDEREÇOS
# ═══════════════════════════════════════════════════════════════════
with tabs[1]:
    st.subheader("2. Endereços")
    st.markdown("**Endereço Fiscal**")
    e1, e2, e3 = st.columns([4, 1, 2])
    with e1:
        end_fiscal_logradouro = st.text_input("Logradouro", key="end_fiscal_logradouro")
    with e2:
        end_fiscal_numero     = st.text_input("Número", key="end_fiscal_numero")
    with e3:
        end_fiscal_complemento = st.text_input("Complemento", key="end_fiscal_complemento")

    e4, e5, e6, e7 = st.columns([3, 2, 1, 2])
    with e4:
        end_fiscal_bairro    = st.text_input("Bairro", key="end_fiscal_bairro")
    with e5:
        end_fiscal_municipio = st.text_input("Município", key="end_fiscal_municipio")
    with e6:
        end_fiscal_estado    = st.text_input("UF", max_chars=2, key="end_fiscal_estado")
    with e7:
        end_fiscal_cep       = st.text_input("CEP", placeholder="00000-000", key="end_fiscal_cep")

    st.markdown("**Endereços de Entrega**")
    entregas = {}
    for n in range(1, 5):
        st.markdown(f"*Entrega {n}*")
        ec1, ec2, ec3 = st.columns([2, 3, 4])
        with ec1:
            entregas[f"end_entrega_{n}_id"] = st.text_input(
                "Identificação da Unidade", key=f"end_entrega_{n}_id")
        with ec2:
            entregas[f"end_entrega_{n}_municipio_estado"] = st.text_input(
                "Município / Estado", key=f"end_entrega_{n}_municipio_estado")
        with ec3:
            entregas[f"end_entrega_{n}_obs"] = st.text_input(
                "Observações", key=f"end_entrega_{n}_obs")

# ═══════════════════════════════════════════════════════════════════
# TAB 3 — CONTATOS
# ═══════════════════════════════════════════════════════════════════
with tabs[2]:
    st.subheader("3. Contatos")

    grupos_contatos = [
        ("Comercial", "com"),
        ("Financeiro", "fin"),
        ("Responsável Técnico", "tec"),
        ("Responsável pela Logística", "log"),
    ]

    contatos = {}
    for label, prefix in grupos_contatos:
        st.markdown(f"**{label}**")
        for idx in (1, 2):
            st.markdown(f"*{idx}°*")
            cc1, cc2, cc3, cc4 = st.columns(4)
            with cc1:
                contatos[f"{prefix}{idx}_nome"]  = st.text_input("Nome",         key=f"{prefix}{idx}_nome")
            with cc2:
                contatos[f"{prefix}{idx}_cargo"] = st.text_input("Cargo",        key=f"{prefix}{idx}_cargo")
            with cc3:
                contatos[f"{prefix}{idx}_tel"]   = st.text_input("Tel/Whatsapp", key=f"{prefix}{idx}_tel")
            with cc4:
                contatos[f"{prefix}{idx}_email"] = st.text_input("E-mail",       key=f"{prefix}{idx}_email")

# ═══════════════════════════════════════════════════════════════════
# TAB 4 — PERFIL & COMERCIAL
# ═══════════════════════════════════════════════════════════════════
with tabs[3]:
    st.subheader("4. Perfil do Cliente")
    p1, p2 = st.columns(2)
    with p1:
        tipo_cliente = st.selectbox("Tipo de cliente", [
            "", "Distribuidor", "Revendedor", "Cooperativa",
            "Produtor Rural", "Importador", "Outro"
        ], key="tipo_cliente")
        area_total_ha   = st.text_input("Área Total (ha)", key="area_total_ha")
        regiao_atuacao  = st.text_input("Região de Atuação", key="regiao_atuacao")
    with p2:
        culturas_principais = st.text_input("Culturas Principais", key="culturas_principais")
        produtos_utilizados = st.text_area("Produtos Biotrop que já compra", height=100, key="produtos_utilizados")

    st.divider()
    st.subheader("5. Condições Comerciais e Financeiras")
    cc1, cc2, cc3 = st.columns(3)
    with cc1:
        condicao_pagamento = st.text_input("Condição de Pagamento", placeholder="ex: 30/60/90 DDL", key="condicao_pagamento")
        limite_credito     = st.text_input("Limite de Crédito (R$)", key="limite_credito")
    with cc2:
        prazo_medio_dias = st.number_input("Prazo Médio (dias)", min_value=0, step=1, value=0, key="prazo_medio_dias")
        forma_pagamento  = st.selectbox("Forma de Pagamento", [
            "", "Boleto", "PIX", "Transferência Bancária", "Cheque", "Cartão", "Misto"
        ], key="forma_pagamento")
    with cc3:
        politica_bonificacao = st.text_area("Política de Bonificação", height=100, key="politica_bonificacao")

    condicoes_especiais = st.text_area("Condições Comerciais Especiais", height=80, key="condicoes_especiais")

# ═══════════════════════════════════════════════════════════════════
# TAB 5 — FATURAMENTO & LOGÍSTICA
# ═══════════════════════════════════════════════════════════════════
with tabs[4]:
    st.subheader("6. Regras de Faturamento")
    f1, f2 = st.columns(2)
    with f1:
        fat_data_limite   = st.text_input("Data limite para faturamento", placeholder="ex: dia 20", key="fat_data_limite")
        fat_exige_po      = sim_nao("Exige pedido formal (PO)?", "fat_exige_po")
        fat_exige_contrato = sim_nao("Exige Contrato?", "fat_exige_contrato")
        fat_conferencia_nf = sim_nao("Exige Conferência Prévia de NF?", "fat_conferencia_nf")
    with f2:
        fat_prazo_nf   = st.text_input("Prazo de Envio para NF", placeholder="ex: 48h antes do embarque", key="fat_prazo_nf")
        fat_shelf_life = st.text_input("Shelf life mínimo do cliente", placeholder="ex: 12 meses", key="fat_shelf_life")
        fat_observacoes = st.text_area("Observações de Faturamento", height=100, key="fat_observacoes")

    st.divider()
    st.subheader("7. Regras Logísticas")
    l1, l2 = st.columns(2)
    with l1:
        log_tipo_entrega    = st.selectbox("Tipo de entrega", ["", "CIF", "FOB", "CIF e FOB"], key="log_tipo_entrega")
        log_agendamento     = sim_nao("Necessita Agendamento?", "log_agendamento")
        log_prazo_agendamento = st.text_input("Prazo Mínimo para Agendamento", placeholder="ex: 48h", key="log_prazo_agendamento")
        log_dias_recebimento  = st.text_input("Dias de Recebimento", placeholder="ex: Segunda a Sexta", key="log_dias_recebimento")
        log_horario_recebimento = st.text_input("Horário de Recebimento", placeholder="ex: 07:00 às 17:00", key="log_horario_recebimento")
        log_aviso_entrega    = sim_nao("Avisar antes da Entrega?", "log_aviso_entrega")
        log_antecedencia_aviso = st.text_input("Antecedência para Aviso (horas)", key="log_antecedencia_aviso")
    with l2:
        log_nf_antecipada    = sim_nao("Exige Envio Antecipado de NF?", "log_nf_antecipada")
        log_antecedencia_nf  = st.text_input("Antecedência de NF (horas)", key="log_antecedencia_nf")
        log_romaneio         = sim_nao("Exige Romaneio?", "log_romaneio")
        log_restricao_transportadora = st.text_input("Restrição de Transportadora", key="log_restricao_transportadora")
        log_regras_acesso    = st.text_area("Regras de Acesso (EPI, docs)", height=80, key="log_regras_acesso")
        log_observacoes      = st.text_area("Observações Logísticas", height=80, key="log_observacoes")

# ═══════════════════════════════════════════════════════════════════
# TAB 6 — LICENÇAS
# ═══════════════════════════════════════════════════════════════════
with tabs[5]:
    st.subheader("8. Licenças e Documentação")
    licencas = [
        ("Licença Ambiental",  "lic_ambiental"),
        ("Licença de Operação", "lic_operacao"),
        ("Registro MAPA",       "lic_mapa"),
        ("Alvará Municipal",    "lic_alvara"),
        ("Certificação ISO",    "lic_iso"),
    ]
    lic_data = {}
    header = st.columns([3, 2, 2, 4])
    header[0].markdown("**Documento**")
    header[1].markdown("**Status**")
    header[2].markdown("**Validade**")
    header[3].markdown("**Observações**")
    for nome_lic, prefix in licencas:
        col_n, col_s, col_v, col_o = st.columns([3, 2, 2, 4])
        with col_n:
            st.markdown(f"<div style='padding-top:8px'>{nome_lic}</div>", unsafe_allow_html=True)
        with col_s:
            lic_data[f"{prefix}_status"]    = status_lic("", f"{prefix}_status")
        with col_v:
            lic_data[f"{prefix}_validade"]  = st.date_input("", value=None, key=f"{prefix}_validade", label_visibility="collapsed")
        with col_o:
            lic_data[f"{prefix}_obs"]       = st.text_input("", key=f"{prefix}_obs", label_visibility="collapsed")

# ═══════════════════════════════════════════════════════════════════
# TAB 7 — ESTRATÉGICO
# ═══════════════════════════════════════════════════════════════════
with tabs[6]:
    st.subheader("9. Informações Estratégicas")
    es1, es2 = st.columns(2)
    with es1:
        estrat_concorrentes     = st.text_area("Concorrentes que o cliente utiliza", height=100, key="estrat_concorrentes")
        estrat_potencial_volume = st.text_input("Potencial de Volume Estimado", placeholder="ex: 500 t/ano", key="estrat_potencial_volume")
        estrat_participacao     = st.text_input("Participação Estimada (%)", placeholder="ex: 30%", key="estrat_participacao")
    with es2:
        estrat_historico        = st.text_area("Histórico de Relacionamentos", height=100, key="estrat_historico")
        estrat_risco            = st.selectbox("Classificação de Risco", ["", "Baixo", "Médio", "Alto"], key="estrat_risco")
        estrat_observacoes      = st.text_area("Observações Estratégicas", height=80, key="estrat_observacoes")

    st.divider()
    st.subheader("10. Complexidade Operacional")
    co1, co2 = st.columns(2)
    with co1:
        compl_classificacao = st.selectbox("Classificação Geral", ["", "Baixa", "Média", "Alta"], key="compl_classificacao")
    with co2:
        compl_justificativa = st.text_area("Justificativa", height=80, key="compl_justificativa")

    st.markdown("**Critérios de Avaliação** (1=baixo · 2=médio · 3=alto)")
    criterios = [
        ("Exigências logísticas",          "compl_log"),
        ("Burocracia de faturamento",       "compl_fat"),
        ("Dificuldade de acesso",           "compl_acesso"),
        ("Nível de exigência operacional",  "compl_operacional"),
    ]
    compl_data = {}
    for nome_crit, prefix in criterios:
        cr1, cr2, cr3 = st.columns([3, 1, 4])
        with cr1:
            st.markdown(f"<div style='padding-top:8px'>{nome_crit}</div>", unsafe_allow_html=True)
        with cr2:
            compl_data[f"{prefix}_nivel"] = nivel("", f"{prefix}_nivel")
        with cr3:
            compl_data[f"{prefix}_obs"]   = st.text_input("", key=f"{prefix}_obs", label_visibility="collapsed", placeholder="Observação")

# ═══════════════════════════════════════════════════════════════════
# TAB 8 — CRÉDITO
# ═══════════════════════════════════════════════════════════════════
with tabs[7]:
    cred_data = {}

    def render_credito(prefix, docs):
        header = st.columns([5, 2, 2, 4])
        header[0].markdown("**Documento**")
        header[1].markdown("**Status**")
        header[2].markdown("**Data de Entrega**")
        header[3].markdown("**Observações**")
        for n, nome_doc in enumerate(docs, start=1):
            c_n, c_s, c_d, c_o = st.columns([5, 2, 2, 4])
            with c_n:
                st.markdown(f"<div style='padding-top:8px;font-size:.9rem'>{nome_doc}</div>", unsafe_allow_html=True)
            with c_s:
                cred_data[f"{prefix}_doc{n}_status"] = status_doc("", f"{prefix}_doc{n}_status")
            with c_d:
                cred_data[f"{prefix}_doc{n}_data"] = st.date_input(
                    "", value=None, key=f"{prefix}_doc{n}_data", label_visibility="collapsed")
            with c_o:
                cred_data[f"{prefix}_doc{n}_obs"] = st.text_input(
                    "", key=f"{prefix}_doc{n}_obs", label_visibility="collapsed")

    st.subheader("11. Documentos para Análise de Crédito — Empresas LTDA")
    render_credito("cred_ltda", [
        "Documentos pessoais dos sócios e cônjuges (RG / CPF)",
        "Comprovante de endereço (sócios)",
        "Certidão de casamento dos sócios (se aplicável)",
        "Contrato social e última alteração contratual",
        "Balanço Patrimonial e DRE — últimos 2 anos (assinados)",
        "Imposto de Renda dos sócios — último exercício",
    ])

    st.divider()
    st.subheader("12. Documentos para Análise de Crédito — Cooperativas / S.A. / Usinas")
    render_credito("cred_coop", [
        "Documentos pessoais dos dirigentes",
        "Comprovante de endereço dos dirigentes",
        "Certidão de casamento dos sócios (se aplicável)",
        "Estatuto Social",
        "Ata de Eleição da Diretoria vigente",
        "Balanço Patrimonial e DRE — últimos 2 anos (assinados)",
    ])

# ═══════════════════════════════════════════════════════════════════
# TAB 9 — CONTROLE INTERNO
# ═══════════════════════════════════════════════════════════════════
with tabs[8]:
    st.subheader("Controle Interno")
    ci1, ci2 = st.columns(2)
    with ci1:
        ctrl_cadastrado_por         = st.text_input("Cadastro realizado por", key="ctrl_cadastrado_por")
        ctrl_data_cadastro          = st.date_input("Data do cadastro", value=date.today(), key="ctrl_data_cadastro")
        ctrl_ultima_atualizacao     = st.date_input("Última atualização", value=date.today(), key="ctrl_ultima_atualizacao")
    with ci2:
        ctrl_responsavel_atualizacao = st.text_input("Responsável pela atualização", key="ctrl_responsavel_atualizacao")
        ctrl_status                  = st.selectbox("Status do cadastro", [
            "", "Em preenchimento", "Completo", "Em revisão", "Aprovado"
        ], key="ctrl_status")

# ═══════════════════════════════════════════════════════════════════
# GERAR FICHA
# ═══════════════════════════════════════════════════════════════════
st.divider()
col_btn, col_info = st.columns([2, 5])

with col_btn:
    gerar = st.button("Gerar Ficha", type="primary", use_container_width=True)

if gerar:
    if not st.session_state.get("razao_social"):
        st.error("Preencha ao menos a Razão Social antes de gerar a ficha.")
    else:
        data = {
            # Bloco 1
            "razao_social":                st.session_state.razao_social,
            "nome_fantasia":               st.session_state.nome_fantasia,
            "cnpj":                        st.session_state.cnpj,
            "inscricao_estadual":          st.session_state.inscricao_estadual,
            "inscricao_municipal":         st.session_state.inscricao_municipal,
            "cnae":                        st.session_state.cnae,
            "data_abertura":               fmt_date(st.session_state.data_abertura),
            "grupo_economico":             st.session_state.grupo_economico,
            "data_inicio_relacionamento":  fmt_date(st.session_state.data_inicio_relacionamento),
            # Responsável Biotrop (hífens para bater com o template)
            "ctrl_responsavel_biotrop":           st.session_state.ctrl_responsavel_biotrop,
            "ctrl_responsavel_biotrop-email":     st.session_state.ctrl_responsavel_biotrop_email,
            "ctrl_responsavel_biotrop-telefone":  st.session_state.ctrl_responsavel_biotrop_tel,
            # Bloco 2 — fiscal
            "end_fiscal_logradouro":  st.session_state.end_fiscal_logradouro,
            "end_fiscal_numero":      st.session_state.end_fiscal_numero,
            "end_fiscal_complemento": st.session_state.end_fiscal_complemento,
            "end_fiscal_bairro":      st.session_state.end_fiscal_bairro,
            "end_fiscal_municipio":   st.session_state.end_fiscal_municipio,
            "end_fiscal_estado":      st.session_state.end_fiscal_estado,
            "end_fiscal_cep":         st.session_state.end_fiscal_cep,
            # Bloco 3 — contatos
            **{k: st.session_state.get(k, "") for k in contatos},
            # Bloco 4
            "tipo_cliente":       st.session_state.tipo_cliente,
            "culturas_principais": st.session_state.culturas_principais,
            "area_total_ha":      st.session_state.area_total_ha,
            "regiao_atuacao":     st.session_state.regiao_atuacao,
            "produtos_utilizados": st.session_state.produtos_utilizados,
            # Bloco 5
            "condicao_pagamento":   st.session_state.condicao_pagamento,
            "prazo_medio_dias":     str(st.session_state.prazo_medio_dias) if st.session_state.prazo_medio_dias else "",
            "limite_credito":       st.session_state.limite_credito,
            "forma_pagamento":      st.session_state.forma_pagamento,
            "politica_bonificacao": st.session_state.politica_bonificacao,
            "condicoes_especiais":  st.session_state.condicoes_especiais,
            # Bloco 6
            "fat_data_limite":    st.session_state.fat_data_limite,
            "fat_exige_po":       st.session_state.fat_exige_po,
            "fat_exige_contrato": st.session_state.fat_exige_contrato,
            "fat_conferencia_nf": st.session_state.fat_conferencia_nf,
            "fat_prazo_nf":       st.session_state.fat_prazo_nf,
            "fat_shelf_life":     st.session_state.fat_shelf_life,
            "fat_observacoes":    st.session_state.fat_observacoes,
            # Bloco 7
            "log_tipo_entrega":            st.session_state.log_tipo_entrega,
            "log_agendamento":             st.session_state.log_agendamento,
            "log_prazo_agendamento":       st.session_state.log_prazo_agendamento,
            "log_dias_recebimento":        st.session_state.log_dias_recebimento,
            "log_horario_recebimento":     st.session_state.log_horario_recebimento,
            "log_aviso_entrega":           st.session_state.log_aviso_entrega,
            "log_antecedencia_aviso":      st.session_state.log_antecedencia_aviso,
            "log_nf_antecipada":           st.session_state.log_nf_antecipada,
            "log_antecedencia_nf":         st.session_state.log_antecedencia_nf,
            "log_romaneio":                st.session_state.log_romaneio,
            "log_restricao_transportadora": st.session_state.log_restricao_transportadora,
            "log_regras_acesso":           st.session_state.log_regras_acesso,
            "log_observacoes":             st.session_state.log_observacoes,
            # Bloco 8 — licenças
            **{k: fmt_date(v) if isinstance(v, (date, datetime)) else (v or "")
               for k, v in lic_data.items()},
            # Bloco 9
            "estrat_concorrentes":     st.session_state.estrat_concorrentes,
            "estrat_potencial_volume": st.session_state.estrat_potencial_volume,
            "estrat_participacao":     st.session_state.estrat_participacao,
            "estrat_historico":        st.session_state.estrat_historico,
            "estrat_risco":            st.session_state.estrat_risco,
            "estrat_observacoes":      st.session_state.estrat_observacoes,
            # Bloco 10
            "compl_classificacao": st.session_state.compl_classificacao,
            "compl_justificativa": st.session_state.compl_justificativa,
            **{k: (v or "") for k, v in compl_data.items()},
            # Blocos 11 e 12
            **{k: fmt_date(v) if isinstance(v, (date, datetime)) else (v or "")
               for k, v in cred_data.items()},
            # Controle Interno
            "ctrl_cadastrado_por":          st.session_state.ctrl_cadastrado_por,
            "ctrl_data_cadastro":           fmt_date(st.session_state.ctrl_data_cadastro),
            "ctrl_ultima_atualizacao":      fmt_date(st.session_state.ctrl_ultima_atualizacao),
            "ctrl_responsavel_atualizacao": st.session_state.ctrl_responsavel_atualizacao,
            "ctrl_status":                  st.session_state.ctrl_status,
        }
        # entregas
        for n in range(1, 5):
            for campo in ("id", "municipio_estado", "obs"):
                k = f"end_entrega_{n}_{campo}"
                data[k] = st.session_state.get(k, "")

        try:
            arquivo = fill_template(data)
            nome_arquivo = f"PRISMA_{razao_social.replace(' ', '_')}.xlsx"
            st.success("Ficha gerada com sucesso!")
            st.download_button(
                label="⬇ Baixar Ficha Excel",
                data=arquivo,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        except FileNotFoundError:
            st.error("Template não encontrado em `template/Prisma - Template.xlsx`. Verifique o repositório.")
        except Exception as e:
            st.error(f"Erro ao gerar ficha: {e}")

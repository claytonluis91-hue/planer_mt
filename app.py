from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from fiscal_core import (
    ParametrosSimulacao,
    SIMPLES_PRIMEIRA_FAIXA_2027,
    carregar_base,
    formatar_documento,
    gerar_excel,
    resumo_simulacao,
    simular_creditos,
    simular_servicos,
)


BASE_DIR = Path(__file__).resolve().parent
ARQUIVOS_DIR = BASE_DIR / "ARQUIVOS"


def moeda(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def percentual(valor: float) -> str:
    return f"{valor:,.4f}%".replace(",", "X").replace(".", ",").replace("X", ".")


@st.cache_data(show_spinner=False)
def carregar() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    return carregar_base(ARQUIVOS_DIR)


st.set_page_config(
    page_title="Projeção IBS e CBS 2027",
    page_icon="🧾",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      .stApp { background: #f6f8f7; }
      [data-testid="stSidebar"] { background: #102a2a; }
      [data-testid="stSidebar"] * { color: #f4f7f6; }
      div[data-testid="stMetric"] {
        background: white; border: 1px solid #dde6e2; border-radius: 14px;
        padding: 16px 18px; box-shadow: 0 2px 10px rgba(16,42,42,.05);
      }
      .hero {
        padding: 24px 28px; border-radius: 18px; margin-bottom: 18px;
        color: white; background: linear-gradient(120deg, #123b3a, #1e6860);
      }
      .hero h1 { margin: 0 0 4px 0; font-size: 2rem; }
      .hero p { margin: 0; opacity: .88; }
      .nota { background:#edf7f2; border-left:4px solid #2f8b72; padding:12px 14px; border-radius:8px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
      <h1>Projeção tributária 2027</h1>
      <p>IBS e CBS sobre serviços médicos, retenções federais e créditos das contas a pagar.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.subheader("Premissas da simulação")
    ibs = st.number_input("IBS padrão (%)", 0.0, 100.0, 0.10, 0.01, format="%.2f")
    cbs = st.number_input("CBS padrão (%)", 0.0, 100.0, 8.50, 0.10, format="%.2f")
    reducao = st.number_input("Redução para saúde (%)", 0.0, 100.0, 60.0, 1.0, format="%.1f")
    st.divider()
    st.caption("Retenções mantidas somente quando já destacadas na nota")
    ir = st.number_input("IRPJ/IRRF (%)", 0.0, 100.0, 1.50, 0.10, format="%.2f")
    csll = st.number_input("CSLL (%)", 0.0, 100.0, 1.00, 0.10, format="%.2f")
    st.divider()
    anexo = st.selectbox(
        "Anexo padrão — fornecedor do Simples",
        SIMPLES_PRIMEIRA_FAIXA_2027["anexo"].tolist(),
        index=2,
        help="O cadastro não informa a atividade/anexo. Use o anexo aplicável ao fornecedor.",
    )
    indefinido_regular = st.toggle(
        "Regime não informado = regular",
        value=False,
        help="Desativado é a opção conservadora: cadastro sem Sim/Não e fornecedor não conciliado não geram crédito.",
    )

parametros = ParametrosSimulacao(
    ibs_pct=ibs,
    cbs_pct=cbs,
    reducao_saude_pct=reducao,
    ir_pct=ir,
    csll_pct=csll,
    anexo_simples_padrao=anexo,
    indefinido_como_regular=indefinido_regular,
)

try:
    with st.spinner("Lendo documentos fiscais e contas a pagar..."):
        servicos_base, contas_base, fornecedores = carregar()
except Exception as erro:
    st.error(f"Não foi possível carregar a pasta ARQUIVOS: {erro}")
    st.stop()

servicos = simular_servicos(servicos_base, parametros)
creditos = simular_creditos(contas_base, parametros)

empresas = sorted(set(servicos["empresa_cnpj"].dropna()) | set(creditos["empresa_cnpj"].dropna()))
rotulos = {cnpj: formatar_documento(cnpj) for cnpj in empresas}
selecionadas = st.multiselect(
    "Estabelecimentos incluídos",
    empresas,
    default=empresas,
    format_func=lambda item: rotulos.get(item, item),
)
servicos_filtrados = servicos[servicos["empresa_cnpj"].isin(selecionadas)].copy()
creditos_filtrados = creditos[creditos["empresa_cnpj"].isin(selecionadas)].copy()
resumo = resumo_simulacao(servicos_filtrados, creditos_filtrados)

aba_resumo, aba_servicos, aba_creditos, aba_conferencia = st.tabs(
    ["Visão geral", "Serviços prestados", "Créditos", "Conferência e premissas"]
)

with aba_resumo:
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Receita analisada", moeda(resumo["receita_bruta"]), f"{len(servicos_filtrados):,} notas".replace(",", "."))
    m2.metric("Débitos IBS + CBS", moeda(resumo["debitos_ibs"] + resumo["debitos_cbs"]))
    m3.metric("Créditos IBS + CBS", moeda(resumo["creditos_ibs"] + resumo["creditos_cbs"]))
    saldo_total = resumo["saldo_ibs"] + resumo["saldo_cbs"]
    m4.metric("Saldo projetado", moeda(saldo_total), "a recolher" if saldo_total >= 0 else "credor")

    st.subheader("Composição do recolhimento")
    composicao = pd.DataFrame(
        {
            "Tributo": ["IBS", "CBS"],
            "Débitos": [resumo["debitos_ibs"], resumo["debitos_cbs"]],
            "Créditos": [resumo["creditos_ibs"], resumo["creditos_cbs"]],
            "Saldo": [resumo["saldo_ibs"], resumo["saldo_cbs"]],
        }
    ).set_index("Tributo")
    st.bar_chart(composicao[["Débitos", "Créditos"]], color=["#1e6860", "#d4a72c"])
    st.dataframe(
        composicao.style.format(moeda),
        width="stretch",
    )

    r1, r2, r3 = st.columns(3)
    r1.metric("Retenções federais nos arquivos", moeda(resumo["retencoes_originais"]))
    r2.metric("Retenções federais em 2027", moeda(resumo["retencoes_2027"]))
    r3.metric("Redução das retenções", moeda(resumo["retencoes_originais"] - resumo["retencoes_2027"]))
    st.markdown(
        f'<div class="nota">Alíquotas efetivas após a redução de saúde: '
        f'IBS <b>{percentual(ibs * (1 - reducao / 100))}</b> e '
        f'CBS <b>{percentual(cbs * (1 - reducao / 100))}</b>.</div>',
        unsafe_allow_html=True,
    )

with aba_servicos:
    formatos = st.multiselect(
        "Formatos de origem",
        sorted(servicos_filtrados["formato"].unique()),
        default=sorted(servicos_filtrados["formato"].unique()),
    )
    mostrar_canceladas = st.toggle("Mostrar notas canceladas/inativas", value=False)
    detalhe_servicos = servicos_filtrados[servicos_filtrados["formato"].isin(formatos)]
    if not mostrar_canceladas:
        detalhe_servicos = detalhe_servicos[detalhe_servicos["considerar"]]
    colunas = [
        "empresa_cnpj_formatado",
        "numero_nota",
        "data_emissao",
        "tomador",
        "valor_servico",
        "base_calculo",
        "debito_ibs_2027",
        "debito_cbs_2027",
        "ir_destacado",
        "csll_destacado",
        "ir_2027",
        "csll_2027",
        "retencao_federal_original",
        "retencao_federal_2027",
        "formato",
    ]
    st.dataframe(
        detalhe_servicos[colunas],
        width="stretch",
        hide_index=True,
        column_config={
            "empresa_cnpj_formatado": "Empresa",
            "numero_nota": "Nota",
            "data_emissao": st.column_config.DateColumn("Emissão", format="DD/MM/YYYY"),
            "tomador": "Tomador",
            "valor_servico": st.column_config.NumberColumn("Serviço", format="R$ %.2f"),
            "base_calculo": st.column_config.NumberColumn("Base IBS/CBS", format="R$ %.2f"),
            "debito_ibs_2027": st.column_config.NumberColumn("IBS", format="R$ %.2f"),
            "debito_cbs_2027": st.column_config.NumberColumn("CBS", format="R$ %.2f"),
            "ir_destacado": "IR no original",
            "csll_destacado": "CSLL no original",
            "ir_2027": st.column_config.NumberColumn("IR 2027", format="R$ %.2f"),
            "csll_2027": st.column_config.NumberColumn("CSLL 2027", format="R$ %.2f"),
            "retencao_federal_original": st.column_config.NumberColumn("Retenção original", format="R$ %.2f"),
            "retencao_federal_2027": st.column_config.NumberColumn("Retenção 2027", format="R$ %.2f"),
            "formato": "Origem",
        },
    )

with aba_creditos:
    c1, c2, c3 = st.columns(3)
    c1.metric("Base das aquisições", moeda(creditos_filtrados["valor_base_credito"].sum()))
    c2.metric("Fornecedores do Simples", f"{creditos_filtrados['simples_nacional'].eq(True).sum():,}".replace(",", "."))
    sem_regra = creditos_filtrados["tratamento_credito"].str.startswith("Sem crédito").sum()
    c3.metric("Lançamentos sem regra", f"{sem_regra:,}".replace(",", "."))
    regimes = st.multiselect(
        "Tratamentos de crédito",
        sorted(creditos_filtrados["tratamento_credito"].unique()),
        default=sorted(creditos_filtrados["tratamento_credito"].unique()),
    )
    detalhe_creditos = creditos_filtrados[creditos_filtrados["tratamento_credito"].isin(regimes)]
    colunas_credito = [
        "empresa_cnpj",
        "codigo_fornecedor",
        "fornecedor",
        "documento",
        "emissao",
        "valor_base_credito",
        "documento_formatado",
        "simples_informado",
        "tratamento_credito",
        "aliquota_credito_ibs_pct",
        "aliquota_credito_cbs_pct",
        "credito_ibs_2027",
        "credito_cbs_2027",
        "situacao",
    ]
    st.dataframe(
        detalhe_creditos[colunas_credito],
        width="stretch",
        hide_index=True,
        column_config={
            "empresa_cnpj": "Empresa",
            "codigo_fornecedor": "Código",
            "fornecedor": "Fornecedor",
            "documento": "Documento",
            "emissao": st.column_config.DateColumn("Emissão", format="DD/MM/YYYY"),
            "valor_base_credito": st.column_config.NumberColumn("Base", format="R$ %.2f"),
            "documento_formatado": "CNPJ/CPF",
            "simples_informado": "Simples",
            "tratamento_credito": "Tratamento",
            "aliquota_credito_ibs_pct": st.column_config.NumberColumn("Alíq. IBS", format="%.4f%%"),
            "aliquota_credito_cbs_pct": st.column_config.NumberColumn("Alíq. CBS", format="%.4f%%"),
            "credito_ibs_2027": st.column_config.NumberColumn("Crédito IBS", format="R$ %.2f"),
            "credito_cbs_2027": st.column_config.NumberColumn("Crédito CBS", format="R$ %.2f"),
            "situacao": "Situação",
        },
    )

with aba_conferencia:
    st.subheader("Primeira faixa do Simples Nacional — 2027/2028")
    tabela_simples = SIMPLES_PRIMEIRA_FAIXA_2027.copy()
    st.dataframe(
        tabela_simples,
        width="stretch",
        hide_index=True,
        column_config={
            "anexo": "Anexo",
            "descricao": "Atividade",
            "aliquota_simples_pct": st.column_config.NumberColumn("Alíquota SN", format="%.2f%%"),
            "partilha_cbs_pct": st.column_config.NumberColumn("Partilha CBS", format="%.2f%%"),
            "partilha_ibs_pct": st.column_config.NumberColumn("Partilha IBS", format="%.2f%%"),
            "credito_cbs_pct": st.column_config.NumberColumn("Crédito CBS efetivo", format="%.5f%%"),
            "credito_ibs_pct": st.column_config.NumberColumn("Crédito IBS efetivo", format="%.5f%%"),
        },
    )
    st.caption(
        "Para optante do Simples 'puro', o crédito do adquirente equivale ao IBS/CBS cobrado no regime. "
        "Se o fornecedor optar pelo regime regular (modelo híbrido), deve ser tratado pela alíquota destacada no documento fiscal."
    )

    st.subheader("Qualidade da conciliação")
    q1, q2, q3 = st.columns(3)
    q1.metric("Notas de serviços lidas", f"{len(servicos_base):,}".replace(",", "."))
    q2.metric("Contas a pagar lidas", f"{len(contas_base):,}".replace(",", "."))
    q3.metric("Fornecedores cadastrados", f"{len(fornecedores):,}".replace(",", "."))
    filtro_pendencia = creditos_filtrados["documento_fornecedor"].isna() | creditos_filtrados[
        "simples_nacional"
    ].isna()
    pendencias = creditos_filtrados.loc[
        filtro_pendencia,
        ["codigo_fornecedor", "fornecedor", "documento_formatado", "simples_informado", "arquivo"],
    ].drop_duplicates()
    if pendencias.empty:
        st.success("Todos os lançamentos selecionados possuem documento e regime do Simples informados.")
    else:
        st.warning("Os itens abaixo não possuem documento ou informação Sim/Não e, por padrão, não geram crédito.")
        st.dataframe(pendencias, width="stretch", hide_index=True)

    with st.expander("Premissas legais e operacionais usadas"):
        st.markdown(
            """
            - Serviços de saúde: redução de 60% das alíquotas, condicionada ao enquadramento do serviço no Anexo III da LC 214/2025.
            - 2027: PIS e Cofins não compõem a retenção projetada; IR e CSLL só são recalculados quando já aparecem destacados no documento original.
            - Aquisições de fornecedor não optante: simulação pela alíquota integral informada no painel.
            - Aquisições do Simples: primeira faixa do anexo selecionado, assumindo recolhimento de IBS/CBS dentro do Simples.
            - Pessoa física, cadastro ausente ou regime indefinido: sem crédito por padrão.
            - O valor da conta a pagar é usado como base estimada; a apropriação definitiva depende do documento fiscal idôneo, destaque e demais requisitos legais.
            """
        )

arquivo_excel = gerar_excel(servicos_filtrados, creditos_filtrados, fornecedores, parametros)
st.download_button(
    "Baixar memória de cálculo em Excel",
    data=arquivo_excel,
    file_name="simulacao_ibs_cbs_2027.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    width="stretch",
)

st.caption(
    "Simulação gerencial com parâmetros editáveis. Confirme classificação, base, direito ao crédito e opção do fornecedor do Simples antes da apuração fiscal."
)

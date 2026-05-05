"""
=============================================================
  ANALISADOR PIS/COFINS MONOFÁSICO – NF-e  |  MVP
  Simples Nacional · Regime de Revenda · CST 04
=============================================================
Dependências:
    pip install streamlit pandas lxml openpyxl

Execução:
    streamlit run app.py
=============================================================
"""

import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO

# ─────────────────────────────────────────────────────────────
#  1. TABELA DE NCM MONOFÁSICOS
#     Baseada na Tabela 4.3.10 da EFD-Contribuições (SPED)
#     ► SUBSTITUA / EXPANDA esta lista conforme necessário
# ─────────────────────────────────────────────────────────────
TABELA_NCM_MONOFASICO: dict[str, str] = {
    # ── Combustíveis e derivados ──────────────────────────────
    "27101112": "Gasolina automotiva comum",
    "27101113": "Gasolina automotiva premium",
    "27101121": "Querosene de aviação",
    "27101131": "Óleo diesel",
    "27101500": "Óleos lubrificantes",
    "27111100": "Gás natural liquefeito (GNL)",
    "27111910": "Gás liquefeito de petróleo (GLP)",

    # ── Farmacêuticos ─────────────────────────────────────────
    "30011000": "Glândulas e outros órgãos para usos opoterápicos",
    "30021000": "Antissoros e imunoglobulinas",
    "30022000": "Vacinas para medicina humana",
    "30023000": "Vacinas para medicina veterinária",
    "30031000": "Medicamentos c/ penicilinas",
    "30032000": "Medicamentos c/ antibióticos",
    "30039099": "Outros medicamentos – mistura não para dose",
    "30041000": "Medicamentos c/ penicilinas (doses)",
    "30042000": "Medicamentos c/ antibióticos (doses)",
    "30043900": "Outros medicamentos hormonais",
    "30049099": "Outros medicamentos para uso humano",

    # ── Cosméticos / Higiene ──────────────────────────────────
    "33011000": "Óleos essenciais de frutas cítricas",
    "33012900": "Outros óleos essenciais",
    "33030010": "Perfumes (extratos)",
    "33030020": "Águas-de-colônia",
    "33041000": "Produtos de maquiagem para lábios",
    "33042000": "Sombras, delineadores",
    "33049900": "Outros produtos de beleza",
    "33051000": "Xampus",
    "33052000": "Preparações para ondulação/alisamento",
    "33053000": "Laquês",
    "33059000": "Outras preparações capilares",
    "33061000": "Dentifrícios",
    "33062000": "Fio dental",
    "33069000": "Outros produtos para higiene bucal",
    "33071000": "Preparações para barbear",
    "33072000": "Desodorantes e antiperspirantes",
    "33074900": "Outros produtos de toucador",

    # ── Bebidas frias ─────────────────────────────────────────
    "22011000": "Água mineral / gaseificada",
    "22019000": "Outras águas",
    "22021000": "Água c/ adição de açúcar / adoçante",
    "22029000": "Outras bebidas não alcoólicas",
    "22030000": "Cerveja de malte",
    "22060000": "Outras bebidas fermentadas",
    "22071000": "Álcool etílico não desnaturado ≥ 80%",
    "22082000": "Aguardente de vinho (conhaque)",
    "22083000": "Uísque",
    "22084000": "Rum e tafia",
    "22085000": "Gim e genebra",
    "22086000": "Vodca",
    "22087000": "Licores",
    "22089900": "Outras bebidas alcoólicas",

    # ── Veículos (automóveis) ─────────────────────────────────
    "87031000": "Veículos para neve, quadriciclos",
    "87032100": "Automóveis cilindrada ≤ 1000 cm³",
    "87032200": "Automóveis 1000 < cil. ≤ 1500 cm³",
    "87032300": "Automóveis 1500 < cil. ≤ 3000 cm³",
    "87032400": "Automóveis cil. > 3000 cm³",
    "87033300": "Automóveis diesel cil. > 2500 cm³",
    "87060010": "Chassis c/ motor para automóveis",
    "87089900": "Outros acessórios para veículos",

    # ── Motos ─────────────────────────────────────────────────
    "87111000": "Motos cilindrada ≤ 50 cm³",
    "87112000": "Motos 50 < cil. ≤ 250 cm³",
    "87113000": "Motos 250 < cil. ≤ 500 cm³",
    "87114000": "Motos 500 < cil. ≤ 800 cm³",
    "87115000": "Motos cil. > 800 cm³",

    # ── Pneus ─────────────────────────────────────────────────
    "40111000": "Pneus novos para automóveis",
    "40112000": "Pneus novos para ônibus/caminhões",
    "40113000": "Pneus novos para aviões",
    "40114000": "Pneus novos para motocicletas",
    "40119100": "Pneus novos – outros",
    "40121100": "Pneus recauchutados para automóveis",
    "40121200": "Pneus recauchutados para ônibus/caminhões",
}


# ─────────────────────────────────────────────────────────────
#  2. LEITURA DE XML – NF-e
# ─────────────────────────────────────────────────────────────
def ler_xml_nfe(conteudo: bytes) -> list[dict]:
    """
    Recebe os bytes de um arquivo XML de NF-e e retorna
    uma lista de dicionários, um por item da nota.
    """
    try:
        root = ET.fromstring(conteudo)
    except ET.ParseError as e:
        raise ValueError(f"XML inválido: {e}")

    # Namespace padrão NF-e 4.0
    ns = {"nfe": "http://www.portalfiscal.inf.br/nfe"}

    itens = []
    # Tenta localizar os elementos <det> (detalhe de item)
    dets = root.findall(".//nfe:det", ns)
    if not dets:
        # Fallback: sem namespace (XMLs simplificados / legados)
        dets = root.findall(".//det")
        ns = {}

    def _find(el, path):
        if ns:
            return el.find(path, ns)
        tag = path.split(":")[-1] if ":" in path else path
        return el.find(f".//{tag}")

    def _text(el, path, default=""):
        node = _find(el, path)
        return node.text.strip() if node is not None and node.text else default

    for det in dets:
        descricao = _text(det, "nfe:prod/nfe:xProd") or _text(det, "prod/xProd")
        ncm_raw   = _text(det, "nfe:prod/nfe:NCM")   or _text(det, "prod/NCM")
        vProd     = _text(det, "nfe:prod/nfe:vProd")  or _text(det, "prod/vProd")

        ncm = ncm_raw.replace(".", "").replace("-", "").strip()

        try:
            valor = float(vProd.replace(",", "."))
        except ValueError:
            valor = 0.0

        itens.append({
            "descricao": descricao or "(sem descrição)",
            "ncm_raw":   ncm_raw,
            "ncm":       ncm,
            "valor":     valor,
        })

    return itens


# ─────────────────────────────────────────────────────────────
#  3. CLASSIFICAÇÃO POR TABELA DE NCM
# ─────────────────────────────────────────────────────────────
def classificar_item(ncm: str, tabela: dict[str, str]) -> tuple[str, str]:
    """
    Retorna (classificacao, motivo).
    A busca é feita do NCM mais específico (8 dígitos)
    ao mais genérico (4 dígitos) para cobrir variações de tabela.
    """
    ncm_limpo = ncm.strip()

    if not ncm_limpo:
        return "INCONSISTÊNCIA", "NCM ausente"
    if not ncm_limpo.isdigit():
        return "INCONSISTÊNCIA", f"NCM inválido: '{ncm_limpo}'"

    # Busca exata (8 dígitos) ou por prefixo (4–7 dígitos)
    for tamanho in (8, 6, 4):
        chave = ncm_limpo[:tamanho].ljust(8, "0")
        if chave in tabela:
            return "MONOFÁSICO", tabela[chave]

    return "NÃO MONOFÁSICO", "NCM fora da tabela"


# ─────────────────────────────────────────────────────────────
#  4. CÁLCULO DE VALORES E RESUMO
# ─────────────────────────────────────────────────────────────
ALIQUOTA_RECUPERACAO = 0.0925  # 9,25% (PIS 1,65% + COFINS 7,60%)

def calcular_resumo(itens_classificados: list[dict]) -> dict:
    """
    Recebe lista de itens já classificados e retorna dicionário
    com totais e estimativa de recuperação tributária.
    """
    total_geral       = sum(i["valor"] for i in itens_classificados)
    total_monofasico  = sum(i["valor"] for i in itens_classificados
                           if i["classificacao"] == "MONOFÁSICO")
    total_nao_mono    = sum(i["valor"] for i in itens_classificados
                           if i["classificacao"] == "NÃO MONOFÁSICO")
    total_inconsist   = sum(i["valor"] for i in itens_classificados
                           if i["classificacao"] == "INCONSISTÊNCIA")
    estimativa_recup  = total_monofasico * ALIQUOTA_RECUPERACAO

    return {
        "total_geral":      total_geral,
        "total_monofasico": total_monofasico,
        "total_nao_mono":   total_nao_mono,
        "total_inconsist":  total_inconsist,
        "estimativa_recup": estimativa_recup,
    }


# ─────────────────────────────────────────────────────────────
#  5. PIPELINE COMPLETO – processa lista de arquivos
# ─────────────────────────────────────────────────────────────
def processar_xmls(arquivos: list[tuple[str, bytes]],
                   tabela: dict[str, str]) -> tuple[list[dict], dict]:
    """
    Recebe [(nome_arquivo, bytes_xml), ...] e a tabela NCM.
    Retorna (lista_itens_classificados, resumo).
    """
    todos_itens = []

    for nome_arquivo, conteudo in arquivos:
        try:
            itens = ler_xml_nfe(conteudo)
        except ValueError as e:
            st.warning(f"⚠️ Erro ao ler '{nome_arquivo}': {e}")
            continue

        for item in itens:
            classificacao, motivo = classificar_item(item["ncm"], tabela)
            todos_itens.append({
                "arquivo":       nome_arquivo,
                "descricao":     item["descricao"],
                "ncm":           item["ncm_raw"],
                "valor":         item["valor"],
                "classificacao": classificacao,
                "motivo":        motivo,
            })

    resumo = calcular_resumo(todos_itens)
    return todos_itens, resumo


# ─────────────────────────────────────────────────────────────
#  6. EXPORTAÇÃO EXCEL
# ─────────────────────────────────────────────────────────────
def gerar_excel(itens: list[dict], resumo: dict) -> bytes:
    """Gera um arquivo Excel com duas abas: Itens e Resumo."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_itens = pd.DataFrame(itens)
        df_itens.columns = ["Arquivo", "Descrição", "NCM",
                            "Valor (R$)", "Classificação", "Motivo"]
        df_itens.to_excel(writer, sheet_name="Itens", index=False)

        df_resumo = pd.DataFrame([{
            "Faturamento Total (R$)":         resumo["total_geral"],
            "Faturamento Monofásico (R$)":    resumo["total_monofasico"],
            "Faturamento Não Monofásico (R$)":resumo["total_nao_mono"],
            "Itens c/ Inconsistência (R$)":   resumo["total_inconsist"],
            "Estimativa de Recuperação (R$)": resumo["estimativa_recup"],
        }])
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)

    return output.getvalue()


# ─────────────────────────────────────────────────────────────
#  7. INTERFACE STREAMLIT
# ─────────────────────────────────────────────────────────────
def formatar_brl(valor: float) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def main():
    st.set_page_config(
        page_title="Analisador PIS/COFINS Monofásico",
        page_icon="📊",
        layout="wide",
    )

    # ── Cabeçalho ────────────────────────────────────────────
    st.title("📊 Analisador PIS/COFINS Monofásico")
    st.caption(
        "Simples Nacional · Regime de Revenda · "
        "Classificação automática por NCM (Tabela 4.3.10 EFD-Contribuições)"
    )
    st.divider()

    # ── Sidebar – configurações ───────────────────────────────
    with st.sidebar:
        st.header("⚙️ Configurações")
        aliquota = st.number_input(
            "Alíquota de recuperação (%)",
            min_value=0.0, max_value=100.0,
            value=ALIQUOTA_RECUPERACAO * 100,
            step=0.05, format="%.2f",
            help="Padrão: 9,25% (PIS 1,65% + COFINS 7,60%)"
        )
        aliquota_decimal = aliquota / 100

        st.markdown("---")
        st.markdown(
            "**Tabela NCM carregada:**  \n"
            f"`{len(TABELA_NCM_MONOFASICO)}` NCMs monofásicos  \n"
            "_Edite `TABELA_NCM_MONOFASICO` no código para atualizar._"
        )

    # ── Upload ────────────────────────────────────────────────
    st.subheader("1️⃣  Upload dos XMLs de NF-e")
    uploaded = st.file_uploader(
        "Selecione um ou mais arquivos XML",
        type=["xml"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.info("📂 Aguardando o upload dos arquivos XML de NF-e…")
        st.stop()

    # ── Processamento ─────────────────────────────────────────
    arquivos = [(f.name, f.read()) for f in uploaded]

    with st.spinner("Processando XMLs…"):
        itens, resumo = processar_xmls(arquivos, TABELA_NCM_MONOFASICO)
        # Recalcula com alíquota customizada
        resumo["estimativa_recup"] = resumo["total_monofasico"] * aliquota_decimal

    if not itens:
        st.error("Nenhum item pôde ser extraído dos arquivos enviados.")
        st.stop()

    df = pd.DataFrame(itens)

    # ── Resumo executivo ──────────────────────────────────────
    st.subheader("2️⃣  Resumo Executivo")
    col1, col2, col3, col4 = st.columns(4)

    col1.metric("💰 Faturamento Total",
                formatar_brl(resumo["total_geral"]))
    col2.metric("🔵 Faturamento Monofásico",
                formatar_brl(resumo["total_monofasico"]))
    col3.metric("⚪ Faturamento Não Monofásico",
                formatar_brl(resumo["total_nao_mono"]))
    col4.metric(f"✅ Estimativa de Recuperação ({aliquota:.2f}%)",
                formatar_brl(resumo["estimativa_recup"]))

    # ── Gráfico de composição ─────────────────────────────────
    st.subheader("3️⃣  Composição do Faturamento")
    graf_data = {
        "Monofásico":     resumo["total_monofasico"],
        "Não Monofásico": resumo["total_nao_mono"],
    }
    if resumo["total_inconsist"] > 0:
        graf_data["Inconsistências"] = resumo["total_inconsist"]

    st.bar_chart(pd.DataFrame.from_dict(
        graf_data, orient="index", columns=["Valor (R$)"]
    ))

    # ── Inconsistências ───────────────────────────────────────
    inconsistencias = df[df["classificacao"] == "INCONSISTÊNCIA"]
    if not inconsistencias.empty:
        st.subheader("⚠️  Inconsistências Encontradas")
        st.warning(
            f"{len(inconsistencias)} item(ns) com NCM ausente ou inválido. "
            "Verifique esses produtos manualmente."
        )
        st.dataframe(
            inconsistencias[["arquivo", "descricao", "ncm", "valor", "motivo"]]
            .rename(columns={
                "arquivo": "Arquivo", "descricao": "Descrição",
                "ncm": "NCM", "valor": "Valor (R$)", "motivo": "Motivo",
            }),
            use_container_width=True,
        )

    # ── Tabela de itens classificados ─────────────────────────
    st.subheader("4️⃣  Itens Classificados")

    # Filtro interativo
    filtro = st.selectbox(
        "Filtrar por classificação",
        ["Todos", "MONOFÁSICO", "NÃO MONOFÁSICO", "INCONSISTÊNCIA"],
    )
    df_exibir = df if filtro == "Todos" else df[df["classificacao"] == filtro]

    def colorir(val):
        cores = {
            "MONOFÁSICO":    "background-color: #d4edda; color: #155724",
            "NÃO MONOFÁSICO":"background-color: #f8f9fa; color: #343a40",
            "INCONSISTÊNCIA":"background-color: #fff3cd; color: #856404",
        }
        return cores.get(val, "")

    st.dataframe(
        df_exibir[["arquivo", "descricao", "ncm",
                   "valor", "classificacao", "motivo"]]
        .rename(columns={
            "arquivo": "Arquivo", "descricao": "Descrição", "ncm": "NCM",
            "valor": "Valor (R$)", "classificacao": "Classificação",
            "motivo": "Motivo / Categoria",
        })
        .style.applymap(colorir, subset=["Classificação"]),
        use_container_width=True,
        height=420,
    )

    # ── Exportar Excel ────────────────────────────────────────
    st.subheader("5️⃣  Exportar Relatório")
    excel_bytes = gerar_excel(itens, resumo)
    st.download_button(
        label="⬇️  Baixar relatório Excel",
        data=excel_bytes,
        file_name="relatorio_pis_cofins_monofasico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── Nota legal ────────────────────────────────────────────
    st.divider()
    st.caption(
        "⚖️ **Aviso:** Este sistema é um MVP para fins de análise preliminar. "
        "A classificação de produtos como monofásicos deve ser validada por "
        "contador ou consultor tributário habilitado. "
        "Os valores de recuperação são estimativas baseadas na legislação vigente."
    )


if __name__ == "__main__":
    main()

"""
=============================================================
  ANALISADOR PIS/COFINS MONOFÁSICO – NF-e  |  MVP v2
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
TABELA_NCM_MONOFASICO: dict = {
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
    "22071000": "Álcool etílico não desnaturado >= 80%",
    "22082000": "Aguardente de vinho (conhaque)",
    "22083000": "Uísque",
    "22084000": "Rum e tafia",
    "22085000": "Gim e genebra",
    "22086000": "Vodca",
    "22087000": "Licores",
    "22089900": "Outras bebidas alcoólicas",
    # ── Veículos ──────────────────────────────────────────────
    "87031000": "Veículos para neve, quadriciclos",
    "87032100": "Automóveis cilindrada <= 1000 cm3",
    "87032200": "Automóveis 1000 < cil. <= 1500 cm3",
    "87032300": "Automóveis 1500 < cil. <= 3000 cm3",
    "87032400": "Automóveis cil. > 3000 cm3",
    "87033300": "Automóveis diesel cil. > 2500 cm3",
    "87060010": "Chassis c/ motor para automóveis",
    "87089900": "Outros acessórios para veículos",
    # ── Motos ─────────────────────────────────────────────────
    "87111000": "Motos cilindrada <= 50 cm3",
    "87112000": "Motos 50 < cil. <= 250 cm3",
    "87113000": "Motos 250 < cil. <= 500 cm3",
    "87114000": "Motos 500 < cil. <= 800 cm3",
    "87115000": "Motos cil. > 800 cm3",
    # ── Pneus ─────────────────────────────────────────────────
    "40111000": "Pneus novos para automóveis",
    "40112000": "Pneus novos para onibus/caminhões",
    "40113000": "Pneus novos para aviões",
    "40114000": "Pneus novos para motocicletas",
    "40119100": "Pneus novos – outros",
    "40121100": "Pneus recauchutados para automóveis",
    "40121200": "Pneus recauchutados para onibus/caminhões",
}


# ─────────────────────────────────────────────────────────────
#  2. LEITURA DE XML – NF-e (versão robusta)
# ─────────────────────────────────────────────────────────────
def ler_xml_nfe(conteudo):
    """
    Recebe os bytes de um arquivo XML de NF-e e retorna
    uma lista de dicionários, um por item da nota.

    Estratégia robusta:
      1. Remove BOM UTF-8 e espaços iniciais
      2. Faz o parse normalmente
      3. Busca <det> em qualquer profundidade via iter()
         (ignora namespace — funciona com NF-e 3.x e 4.x,
          nfeProc, enviNFe, com ou sem namespace)
      4. Extrai campos pela tag local
    """
    # Remove BOM UTF-8 (EF BB BF) e espaços antes da declaração XML
    if isinstance(conteudo, bytes):
        conteudo = conteudo.lstrip(b"\xef\xbb\xbf").strip()

    try:
        root = ET.fromstring(conteudo)
    except ET.ParseError as e:
        raise ValueError(f"Erro ao interpretar XML: {e}")

    # Helper: retorna a tag local sem namespace
    def local(node):
        t = node.tag
        return t.split("}")[-1] if "}" in t else t

    # Helper: busca primeiro descendente com tag local
    def find_local(el, tag):
        for node in el.iter():
            if local(node) == tag:
                return node
        return None

    # Helper: texto de um descendente com tag local
    def text_local(el, tag, default=""):
        node = find_local(el, tag)
        return node.text.strip() if node is not None and node.text else default

    # Localiza todos os <det> na árvore inteira
    dets = [n for n in root.iter() if local(n) == "det"]

    itens = []
    for det in dets:
        prod = find_local(det, "prod")
        if prod is None:
            continue

        descricao = text_local(prod, "xProd")
        ncm_raw   = text_local(prod, "NCM")
        vprod_str = text_local(prod, "vProd")

        ncm = ncm_raw.replace(".", "").replace("-", "").strip()

        try:
            valor = float(vprod_str.replace(",", "."))
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
def classificar_item(ncm, tabela):
    """
    Retorna (classificacao, motivo).
    Busca do NCM mais específico (8 dígitos) ao mais genérico (4 dígitos).
    """
    ncm_limpo = ncm.strip()

    if not ncm_limpo:
        return "INCONSISTENCIA", "NCM ausente"
    if not ncm_limpo.isdigit():
        return "INCONSISTENCIA", "NCM invalido: '{}'".format(ncm_limpo)

    for tamanho in (8, 6, 4):
        chave = ncm_limpo[:tamanho].ljust(8, "0")
        if chave in tabela:
            return "MONOFASICO", tabela[chave]

    return "NAO MONOFASICO", "NCM fora da tabela"


# ─────────────────────────────────────────────────────────────
#  4. CÁLCULO DE VALORES E RESUMO
# ─────────────────────────────────────────────────────────────
ALIQUOTA_RECUPERACAO = 0.0925  # 9,25% (PIS 1,65% + COFINS 7,60%)

def calcular_resumo(itens_classificados, aliquota=ALIQUOTA_RECUPERACAO):
    total_geral      = sum(i["valor"] for i in itens_classificados)
    total_mono       = sum(i["valor"] for i in itens_classificados if i["classificacao"] == "MONOFASICO")
    total_nao_mono   = sum(i["valor"] for i in itens_classificados if i["classificacao"] == "NAO MONOFASICO")
    total_inconsist  = sum(i["valor"] for i in itens_classificados if i["classificacao"] == "INCONSISTENCIA")
    estimativa_recup = total_mono * aliquota
    return {
        "total_geral":      total_geral,
        "total_monofasico": total_mono,
        "total_nao_mono":   total_nao_mono,
        "total_inconsist":  total_inconsist,
        "estimativa_recup": estimativa_recup,
    }


# ─────────────────────────────────────────────────────────────
#  5. PIPELINE COMPLETO
# ─────────────────────────────────────────────────────────────
def processar_xmls(arquivos, tabela, aliquota=ALIQUOTA_RECUPERACAO):
    """
    arquivos: lista de (nome_arquivo, bytes_xml)
    Retorna (lista_itens_classificados, resumo)
    """
    todos_itens = []
    for nome, conteudo in arquivos:
        try:
            itens = ler_xml_nfe(conteudo)
        except ValueError as e:
            st.warning("Erro ao ler '{}': {}".format(nome, e))
            continue

        for item in itens:
            classif, motivo = classificar_item(item["ncm"], tabela)
            todos_itens.append({
                "arquivo":       nome,
                "descricao":     item["descricao"],
                "ncm":           item["ncm_raw"],
                "valor":         item["valor"],
                "classificacao": classif,
                "motivo":        motivo,
            })

    resumo = calcular_resumo(todos_itens, aliquota)
    return todos_itens, resumo


# ─────────────────────────────────────────────────────────────
#  6. EXPORTAÇÃO EXCEL
# ─────────────────────────────────────────────────────────────
def gerar_excel(itens, resumo):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_itens = pd.DataFrame(itens)
        df_itens.columns = ["Arquivo", "Descricao", "NCM", "Valor (R$)", "Classificacao", "Motivo"]
        df_itens.to_excel(writer, sheet_name="Itens", index=False)

        df_resumo = pd.DataFrame([{
            "Faturamento Total (R$)":          resumo["total_geral"],
            "Faturamento Monofasico (R$)":     resumo["total_monofasico"],
            "Faturamento Nao Monofasico (R$)": resumo["total_nao_mono"],
            "Itens c/ Inconsistencia (R$)":    resumo["total_inconsist"],
            "Estimativa de Recuperacao (R$)":  resumo["estimativa_recup"],
        }])
        df_resumo.to_excel(writer, sheet_name="Resumo", index=False)
    return output.getvalue()


# ─────────────────────────────────────────────────────────────
#  7. INTERFACE STREAMLIT
# ─────────────────────────────────────────────────────────────
def formatar_brl(valor):
    return "R$ {:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")


def main():
    st.set_page_config(
        page_title="Analisador PIS/COFINS Monofasico",
        page_icon="📊",
        layout="wide",
    )

    st.title("📊 Analisador PIS/COFINS Monofasico")
    st.caption(
        "Simples Nacional · Regime de Revenda · "
        "Classificacao automatica por NCM (Tabela 4.3.10 EFD-Contribuicoes)"
    )
    st.divider()

    # ── Sidebar ───────────────────────────────────────────────
    with st.sidebar:
        st.header("Configuracoes")
        aliquota = st.number_input(
            "Aliquota de recuperacao (%)",
            min_value=0.0, max_value=100.0,
            value=9.25, step=0.05, format="%.2f",
            help="Padrao: 9,25% (PIS 1,65% + COFINS 7,60%)"
        )
        aliquota_decimal = aliquota / 100
        st.markdown("---")
        st.markdown(
            "**Tabela NCM carregada:**  \n"
            "`{}` NCMs monofasicos  \n"
            "_Edite `TABELA_NCM_MONOFASICO` no codigo para atualizar._".format(
                len(TABELA_NCM_MONOFASICO)
            )
        )

    # ── Upload ────────────────────────────────────────────────
    st.subheader("1  Upload dos XMLs de NF-e")
    uploaded = st.file_uploader(
        "Selecione um ou mais arquivos XML",
        type=["xml"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.info("Aguardando o upload dos arquivos XML de NF-e...")
        st.stop()

    arquivos = [(f.name, f.read()) for f in uploaded]

    with st.spinner("Processando XMLs..."):
        itens, resumo = processar_xmls(arquivos, TABELA_NCM_MONOFASICO, aliquota_decimal)

    if not itens:
        st.error("Nenhum item pode ser extraido dos arquivos enviados.")
        st.stop()

    df = pd.DataFrame(itens)

    # ── Métricas ──────────────────────────────────────────────
    st.subheader("2  Resumo Executivo")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Faturamento Total",         formatar_brl(resumo["total_geral"]))
    col2.metric("Faturamento Monofasico",    formatar_brl(resumo["total_monofasico"]))
    col3.metric("Faturamento Nao Monofasico",formatar_brl(resumo["total_nao_mono"]))
    col4.metric(
        "Estimativa de Recuperacao ({:.2f}%)".format(aliquota),
        formatar_brl(resumo["estimativa_recup"])
    )

    # ── Gráfico ───────────────────────────────────────────────
    st.subheader("3  Composicao do Faturamento")
    graf = {"Monofasico": resumo["total_monofasico"], "Nao Monofasico": resumo["total_nao_mono"]}
    if resumo["total_inconsist"] > 0:
        graf["Inconsistencias"] = resumo["total_inconsist"]
    st.bar_chart(pd.DataFrame.from_dict(graf, orient="index", columns=["Valor (R$)"]))

    # ── Inconsistências ───────────────────────────────────────
    inconsist = df[df["classificacao"] == "INCONSISTENCIA"]
    if not inconsist.empty:
        st.subheader("Inconsistencias Encontradas")
        st.warning("{} item(ns) com NCM ausente ou invalido. Verifique manualmente.".format(len(inconsist)))
        st.dataframe(
            inconsist[["arquivo", "descricao", "ncm", "valor", "motivo"]].rename(columns={
                "arquivo": "Arquivo", "descricao": "Descricao",
                "ncm": "NCM", "valor": "Valor (R$)", "motivo": "Motivo"
            }),
            use_container_width=True,
        )

    # ── Tabela de itens ───────────────────────────────────────
    st.subheader("4  Itens Classificados")
    filtro = st.selectbox(
        "Filtrar por classificacao",
        ["Todos", "MONOFASICO", "NAO MONOFASICO", "INCONSISTENCIA"],
    )
    df_exib = df if filtro == "Todos" else df[df["classificacao"] == filtro]

    st.dataframe(
        df_exib[["arquivo", "descricao", "ncm", "valor", "classificacao", "motivo"]].rename(columns={
            "arquivo": "Arquivo", "descricao": "Descricao", "ncm": "NCM",
            "valor": "Valor (R$)", "classificacao": "Classificacao", "motivo": "Motivo"
        }),
        use_container_width=True,
        height=420,
    )

    # ── Export ────────────────────────────────────────────────
    st.subheader("5  Exportar Relatorio")
    excel_bytes = gerar_excel(itens, resumo)
    st.download_button(
        label="Baixar relatorio Excel",
        data=excel_bytes,
        file_name="relatorio_pis_cofins_monofasico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.caption(
        "Aviso: Este sistema e um MVP para fins de analise preliminar. "
        "A classificacao de produtos como monofasicos deve ser validada por "
        "contador ou consultor tributario habilitado."
    )


if __name__ == "__main__":
    main()

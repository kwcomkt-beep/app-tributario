"""
=============================================================
  ANALISADOR PIS/COFINS MONOFASICO + APURACAO PGDAS-D
  Simples Nacional · Regime de Revenda · CST 04  |  MVP v3
=============================================================
Dependencias:
    pip install streamlit pandas openpyxl

Execucao:
    streamlit run app.py
=============================================================
"""

import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO, StringIO
from datetime import datetime

# ─────────────────────────────────────────────────────────────
#  1. TABELA DE NCM MONOFASICOS
#     Baseada na Tabela 4.3.10 da EFD-Contribuicoes (SPED)
#     Edite este dicionario para adicionar/remover NCMs
# ─────────────────────────────────────────────────────────────
TABELA_NCM_MONOFASICO = {
    "27101112": "Gasolina automotiva comum",
    "27101113": "Gasolina automotiva premium",
    "27101121": "Querosene de aviacao",
    "27101131": "Oleo diesel",
    "27101500": "Oleos lubrificantes",
    "27111100": "Gas natural liquefeito (GNL)",
    "27111910": "Gas liquefeito de petroleo (GLP)",
    "30011000": "Glandulas e outros orgaos para usos opoterapicos",
    "30021000": "Antissoros e imunoglobulinas",
    "30022000": "Vacinas para medicina humana",
    "30023000": "Vacinas para medicina veterinaria",
    "30031000": "Medicamentos c/ penicilinas",
    "30032000": "Medicamentos c/ antibioticos",
    "30039099": "Outros medicamentos – mistura nao para dose",
    "30041000": "Medicamentos c/ penicilinas (doses)",
    "30042000": "Medicamentos c/ antibioticos (doses)",
    "30043900": "Outros medicamentos hormonais",
    "30049099": "Outros medicamentos para uso humano",
    "33011000": "Oleos essenciais de frutas citricas",
    "33012900": "Outros oleos essenciais",
    "33030010": "Perfumes (extratos)",
    "33030020": "Aguas-de-colonia",
    "33041000": "Produtos de maquiagem para labios",
    "33042000": "Sombras, delineadores",
    "33049900": "Outros produtos de beleza",
    "33051000": "Xampus",
    "33052000": "Preparacoes para ondulacao/alisamento",
    "33053000": "Laques",
    "33059000": "Outras preparacoes capilares",
    "33061000": "Dentifricio",
    "33062000": "Fio dental",
    "33069000": "Outros produtos para higiene bucal",
    "33071000": "Preparacoes para barbear",
    "33072000": "Desodorantes e antiperspirantes",
    "33074900": "Outros produtos de toucador",
    "22011000": "Agua mineral / gaseificada",
    "22019000": "Outras aguas",
    "22021000": "Agua c/ adicao de acucar / adocante",
    "22029000": "Outras bebidas nao alcoolicas",
    "22030000": "Cerveja de malte",
    "22060000": "Outras bebidas fermentadas",
    "22071000": "Alcool etilico nao desnaturado >= 80%",
    "22082000": "Aguardente de vinho (conhaque)",
    "22083000": "Uisque",
    "22084000": "Rum e tafia",
    "22085000": "Gim e genebra",
    "22086000": "Vodca",
    "22087000": "Licores",
    "22089900": "Outras bebidas alcoolicas",
    "87031000": "Veiculos para neve, quadriciclos",
    "87032100": "Automoveis cilindrada <= 1000 cm3",
    "87032200": "Automoveis 1000 < cil. <= 1500 cm3",
    "87032300": "Automoveis 1500 < cil. <= 3000 cm3",
    "87032400": "Automoveis cil. > 3000 cm3",
    "87033300": "Automoveis diesel cil. > 2500 cm3",
    "87060010": "Chassis c/ motor para automoveis",
    "87089900": "Outros acessorios para veiculos",
    "87111000": "Motos cilindrada <= 50 cm3",
    "87112000": "Motos 50 < cil. <= 250 cm3",
    "87113000": "Motos 250 < cil. <= 500 cm3",
    "87114000": "Motos 500 < cil. <= 800 cm3",
    "87115000": "Motos cil. > 800 cm3",
    "40111000": "Pneus novos para automoveis",
    "40112000": "Pneus novos para onibus/caminhoes",
    "40113000": "Pneus novos para avioes",
    "40114000": "Pneus novos para motocicletas",
    "40119100": "Pneus novos – outros",
    "40121100": "Pneus recauchutados para automoveis",
    "40121200": "Pneus recauchutados para onibus/caminhoes",
}

# ─────────────────────────────────────────────────────────────
#  2. TABELAS DO SIMPLES NACIONAL
#     Fonte: LC 123/2006 – Resolucao CGSN 140/2018
#     Estrutura: (limite_superior, aliq_nominal, parcela_deduzir)
# ─────────────────────────────────────────────────────────────
TABELAS_SIMPLES = {
    "Anexo I – Comercio": [
        (180_000,       0.04,   0.0),
        (360_000,       0.073,  5_940.0),
        (720_000,       0.095,  13_860.0),
        (1_800_000,     0.107,  22_500.0),
        (3_600_000,     0.143,  87_300.0),
        (4_800_000,     0.19,   378_000.0),
    ],
    "Anexo II – Industria": [
        (180_000,       0.045,  0.0),
        (360_000,       0.078,  5_940.0),
        (720_000,       0.10,   13_860.0),
        (1_800_000,     0.113,  22_500.0),
        (3_600_000,     0.147,  85_500.0),
        (4_800_000,     0.30,   720_000.0),
    ],
    "Anexo III – Servicos A": [
        (180_000,       0.06,   0.0),
        (360_000,       0.112,  9_360.0),
        (720_000,       0.135,  17_640.0),
        (1_800_000,     0.16,   35_640.0),
        (3_600_000,     0.21,   125_640.0),
        (4_800_000,     0.33,   648_000.0),
    ],
    "Anexo IV – Servicos B": [
        (180_000,       0.045,  0.0),
        (360_000,       0.09,   8_100.0),
        (720_000,       0.102,  12_420.0),
        (1_800_000,     0.14,   39_780.0),
        (3_600_000,     0.22,   183_780.0),
        (4_800_000,     0.33,   828_000.0),
    ],
    "Anexo V – Servicos C": [
        (180_000,       0.15,   0.0),
        (360_000,       0.18,   5_400.0),
        (720_000,       0.195,  13_500.0),
        (1_800_000,     0.205,  20_700.0),
        (3_600_000,     0.23,   62_100.0),
        (4_800_000,     0.305,  540_000.0),
    ],
}

ALIQUOTA_PIS_COFINS = 0.0925  # 9,25% – estimativa PIS + COFINS monofasico


# ─────────────────────────────────────────────────────────────
#  3. LEITURA DE XML – NF-e (robusto: ignora namespace)
# ─────────────────────────────────────────────────────────────
def local_tag(node):
    t = node.tag
    return t.split("}")[-1] if "}" in t else t

def find_local(el, tag):
    for node in el.iter():
        if local_tag(node) == tag:
            return node
    return None

def text_local(el, tag, default=""):
    node = find_local(el, tag)
    return node.text.strip() if node is not None and node.text else default


def extrair_data(root):
    """
    Extrai dhEmi do XML e retorna string 'YYYY-MM'.
    Fallback para dEmi (NF-e mais antigas).
    Retorna 'SEM-DATA' se nao encontrar.
    """
    dh = text_local(root, "dhEmi") or text_local(root, "dEmi")
    if not dh:
        return "SEM-DATA"
    # dhEmi pode vir como "2024-03-15T10:30:00-03:00" ou "2024-03-15"
    try:
        return dh[:7]  # "YYYY-MM"
    except Exception:
        return "SEM-DATA"


def ler_xml_nfe(conteudo):
    """
    Recebe bytes de um XML de NF-e.
    Retorna (lista_de_itens, mes_ano_str).
    Cada item: {descricao, ncm_raw, ncm, valor}
    """
    if isinstance(conteudo, bytes):
        conteudo = conteudo.lstrip(b"\xef\xbb\xbf").strip()

    try:
        root = ET.fromstring(conteudo)
    except ET.ParseError as e:
        raise ValueError("Erro ao interpretar XML: {}".format(e))

    mes_ano = extrair_data(root)

    dets = [n for n in root.iter() if local_tag(n) == "det"]
    itens = []
    for det in dets:
        prod = find_local(det, "prod")
        if prod is None:
            continue

        descricao = text_local(prod, "xProd")
        ncm_raw   = text_local(prod, "NCM")
        vprod_str = text_local(prod, "vProd")
        ncm       = ncm_raw.replace(".", "").replace("-", "").strip()

        try:
            valor = float(vprod_str.replace(",", "."))
        except ValueError:
            valor = 0.0

        itens.append({
            "descricao": descricao or "(sem descricao)",
            "ncm_raw":   ncm_raw,
            "ncm":       ncm,
            "valor":     valor,
        })

    return itens, mes_ano


# ─────────────────────────────────────────────────────────────
#  4. CLASSIFICACAO POR TABELA NCM
# ─────────────────────────────────────────────────────────────
def classificar_item(ncm, tabela):
    """
    Retorna (classificacao, motivo).
    Busca por 8, 6 ou 4 digitos.
    """
    ncm_limpo = ncm.strip()
    if not ncm_limpo:
        return "INCONSISTENCIA", "NCM ausente"
    if not ncm_limpo.isdigit():
        return "INCONSISTENCIA", "NCM invalido: {}".format(ncm_limpo)
    for tam in (8, 6, 4):
        chave = ncm_limpo[:tam].ljust(8, "0")
        if chave in tabela:
            return "MONOFASICO", tabela[chave]
    return "NAO MONOFASICO", "NCM fora da tabela"


# ─────────────────────────────────────────────────────────────
#  5. AGRUPAMENTO MENSAL
# ─────────────────────────────────────────────────────────────
def agrupar_por_mes(itens_classificados):
    """
    Agrupa lista de itens classificados por mes.
    Retorna lista de dicts:
      {mes, receita_total, receita_monofasica, receita_tributavel}
    """
    meses = {}
    for item in itens_classificados:
        m = item.get("mes", "SEM-DATA")
        if m not in meses:
            meses[m] = {"receita_total": 0.0, "receita_monofasica": 0.0}
        meses[m]["receita_total"] += item["valor"]
        if item["classificacao"] == "MONOFASICO":
            meses[m]["receita_monofasica"] += item["valor"]

    resultado = []
    for mes in sorted(meses.keys()):
        total = meses[mes]["receita_total"]
        mono  = meses[mes]["receita_monofasica"]
        resultado.append({
            "mes":               mes,
            "receita_total":     total,
            "receita_monofasica":mono,
            "receita_tributavel":total - mono,
        })
    return resultado


# ─────────────────────────────────────────────────────────────
#  6. CALCULO DA ALIQUOTA EFETIVA (PGDAS-D)
# ─────────────────────────────────────────────────────────────
def calcular_aliquota_efetiva(rbt12, anexo):
    """
    Calcula a aliquota efetiva do Simples Nacional.
    Formula: (RBT12 * aliq_nominal - parcela_deduzir) / RBT12
    Retorna (aliquota_efetiva, aliquota_nominal, parcela_deduzir, faixa_idx)
    """
    faixas = TABELAS_SIMPLES.get(anexo, [])
    if not faixas:
        return 0.0, 0.0, 0.0, -1

    if rbt12 <= 0:
        return 0.0, 0.0, 0.0, 0

    for idx, (limite, aliq_nom, parcela) in enumerate(faixas):
        if rbt12 <= limite:
            efetiva = (rbt12 * aliq_nom - parcela) / rbt12
            return round(efetiva, 6), aliq_nom, parcela, idx + 1

    # Acima do limite maximo do Simples
    return None, None, None, -1


# ─────────────────────────────────────────────────────────────
#  7. CALCULO DO DAS
# ─────────────────────────────────────────────────────────────
def calcular_das(agrupamento_mensal, aliquota_efetiva):
    """
    Para cada mes calcula:
      DAS pago    = receita_total * aliquota_efetiva
      DAS correto = receita_tributavel * aliquota_efetiva
    Retorna lista de dicts com colunas de DAS.
    """
    resultado = []
    for row in agrupamento_mensal:
        das_pago    = row["receita_total"]     * aliquota_efetiva
        das_correto = row["receita_tributavel"] * aliquota_efetiva
        resultado.append({**row, "das_pago": das_pago, "das_correto": das_correto})
    return resultado


# ─────────────────────────────────────────────────────────────
#  8. CALCULO DO CREDITO
# ─────────────────────────────────────────────────────────────
def calcular_credito(dados_das):
    """
    Credito_mensal = DAS_pago - DAS_correto
    Retorna lista com coluna 'credito' e o total acumulado.
    """
    resultado = []
    total_credito = 0.0
    for row in dados_das:
        credito = row["das_pago"] - row["das_correto"]
        total_credito += credito
        resultado.append({**row, "credito": credito})
    return resultado, total_credito


# ─────────────────────────────────────────────────────────────
#  9. CALCULO DE RESUMO GERAL (mantido do v2)
# ─────────────────────────────────────────────────────────────
ALIQUOTA_RECUPERACAO = ALIQUOTA_PIS_COFINS

def calcular_resumo(itens_classificados, aliquota=ALIQUOTA_RECUPERACAO):
    total_geral     = sum(i["valor"] for i in itens_classificados)
    total_mono      = sum(i["valor"] for i in itens_classificados if i["classificacao"] == "MONOFASICO")
    total_nao_mono  = sum(i["valor"] for i in itens_classificados if i["classificacao"] == "NAO MONOFASICO")
    total_inconsist = sum(i["valor"] for i in itens_classificados if i["classificacao"] == "INCONSISTENCIA")
    estimativa      = total_mono * aliquota
    return {
        "total_geral":      total_geral,
        "total_monofasico": total_mono,
        "total_nao_mono":   total_nao_mono,
        "total_inconsist":  total_inconsist,
        "estimativa_recup": estimativa,
    }


# ─────────────────────────────────────────────────────────────
#  10. PIPELINE COMPLETO
# ─────────────────────────────────────────────────────────────
def processar_xmls(arquivos, tabela, aliquota=ALIQUOTA_RECUPERACAO):
    todos_itens = []
    for nome, conteudo in arquivos:
        try:
            itens, mes_ano = ler_xml_nfe(conteudo)
        except ValueError as e:
            st.warning("Erro ao ler '{}': {}".format(nome, e))
            continue

        for item in itens:
            classif, motivo = classificar_item(item["ncm"], tabela)
            todos_itens.append({
                "arquivo":       nome,
                "mes":           mes_ano,
                "descricao":     item["descricao"],
                "ncm":           item["ncm_raw"],
                "valor":         item["valor"],
                "classificacao": classif,
                "motivo":        motivo,
            })

    resumo = calcular_resumo(todos_itens, aliquota)
    return todos_itens, resumo


# ─────────────────────────────────────────────────────────────
#  11. EXPORTACOES
# ─────────────────────────────────────────────────────────────
def gerar_excel(itens, resumo):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_i = pd.DataFrame(itens)
        df_i.columns = ["Arquivo", "Mes", "Descricao", "NCM",
                        "Valor (R$)", "Classificacao", "Motivo"]
        df_i.to_excel(writer, sheet_name="Itens", index=False)

        df_r = pd.DataFrame([{
            "Faturamento Total (R$)":         resumo["total_geral"],
            "Faturamento Monofasico (R$)":    resumo["total_monofasico"],
            "Faturamento Nao Monofasico (R$)":resumo["total_nao_mono"],
            "Inconsistencias (R$)":           resumo["total_inconsist"],
            "Estimativa Recuperacao (R$)":    resumo["estimativa_recup"],
        }])
        df_r.to_excel(writer, sheet_name="Resumo", index=False)
    return output.getvalue()


def gerar_csv(dados_apuracao):
    df = pd.DataFrame(dados_apuracao)
    df.columns = ["Mes", "Receita Total", "Rec. Monofasica",
                  "Rec. Tributavel", "DAS Pago", "DAS Correto", "Credito"]
    return df.to_csv(index=False, sep=";", decimal=",").encode("utf-8")


# ─────────────────────────────────────────────────────────────
#  12. HELPERS DE FORMATACAO
# ─────────────────────────────────────────────────────────────
def brl(v):
    return "R$ {:,.2f}".format(v).replace(",", "X").replace(".", ",").replace("X", ".")

def pct(v):
    return "{:.4f}%".format(v * 100)

def fmt_mes(m):
    """Converte 'YYYY-MM' para 'MMM/YYYY' (ex: 'mar/2024')."""
    try:
        return datetime.strptime(m, "%Y-%m").strftime("%b/%Y")
    except Exception:
        return m


# ─────────────────────────────────────────────────────────────
#  13. INTERFACE STREAMLIT
# ─────────────────────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="PIS/COFINS + PGDAS-D",
        page_icon="📊",
        layout="wide",
    )

    st.title("📊 Analisador PIS/COFINS Monofasico + Apuracao PGDAS-D")
    st.caption("Simples Nacional · Regime de Revenda · Classificacao por NCM (Tabela 4.3.10 EFD-Contribuicoes)")
    st.divider()

    # ── SIDEBAR ───────────────────────────────────────────────
    with st.sidebar:
        st.header("Configuracoes")

        st.subheader("Simples Nacional")
        anexo = st.selectbox(
            "Anexo",
            list(TABELAS_SIMPLES.keys()),
            help="Selecione o anexo da sua atividade"
        )
        rbt12 = st.number_input(
            "RBT12 – Receita Bruta 12 meses (R$)",
            min_value=0.0,
            value=360_000.0,
            step=1_000.0,
            format="%.2f",
            help="Receita bruta acumulada dos ultimos 12 meses antes do periodo apurado"
        )

        st.markdown("---")
        st.subheader("PIS/COFINS")
        aliquota_pc = st.number_input(
            "Aliquota estimativa PIS+COFINS (%)",
            min_value=0.0, max_value=100.0,
            value=9.25, step=0.05, format="%.2f",
            help="Usada no calculo da estimativa simples de credito"
        )
        aliquota_decimal = aliquota_pc / 100

        st.markdown("---")
        st.markdown(
            "**Tabela NCM carregada:**  \n"
            "`{}` NCMs monofasicos".format(len(TABELA_NCM_MONOFASICO))
        )

        # Preview aliquota efetiva
        if rbt12 > 0:
            ef, nom, ded, faixa = calcular_aliquota_efetiva(rbt12, anexo)
            st.markdown("---")
            st.subheader("Preview Aliquota")
            if ef is None:
                st.error("RBT12 acima do limite do Simples (R$ 4,8M)")
            else:
                st.metric("Aliquota Efetiva", pct(ef))
                st.caption("Faixa {} · Nominal {} · Deducao {}".format(
                    faixa, pct(nom), brl(ded)))

    # ── UPLOAD ────────────────────────────────────────────────
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

    # ── METRICAS GERAIS ───────────────────────────────────────
    st.subheader("2  Resumo Geral")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Faturamento Total",          brl(resumo["total_geral"]))
    c2.metric("Faturamento Monofasico",     brl(resumo["total_monofasico"]))
    c3.metric("Faturamento Nao Monofasico", brl(resumo["total_nao_mono"]))
    c4.metric(
        "Estimativa PIS+COFINS ({:.2f}%)".format(aliquota_pc),
        brl(resumo["estimativa_recup"])
    )

    # ── GRAFICO ───────────────────────────────────────────────
    st.subheader("3  Composicao do Faturamento")
    graf = {
        "Monofasico":    resumo["total_monofasico"],
        "Nao Monofasico":resumo["total_nao_mono"],
    }
    if resumo["total_inconsist"] > 0:
        graf["Inconsistencias"] = resumo["total_inconsist"]
    st.bar_chart(pd.DataFrame.from_dict(graf, orient="index", columns=["Valor (R$)"]))

    # ── APURACAO PGDAS-D ──────────────────────────────────────
    st.subheader("4  Apuracao PGDAS-D – Credito por Monofasico")

    ef, nom, ded, faixa = calcular_aliquota_efetiva(rbt12, anexo)

    if ef is None:
        st.error(
            "RBT12 informado ({}}) excede o limite maximo do Simples Nacional "
            "(R$ 4.800.000,00). Verifique o valor informado.".format(brl(rbt12))
        )
    elif rbt12 == 0:
        st.warning("Informe o RBT12 na sidebar para calcular a apuracao.")
    else:
        # Pipeline de apuracao
        agrupado   = agrupar_por_mes(itens)
        com_das    = calcular_das(agrupado, ef)
        apuracao, total_credito = calcular_credito(com_das)

        # Tabela principal
        linhas = []
        for row in apuracao:
            linhas.append({
                "Mes":             fmt_mes(row["mes"]),
                "Receita Total":   brl(row["receita_total"]),
                "Monofasico":      brl(row["receita_monofasica"]),
                "Tributavel":      brl(row["receita_tributavel"]),
                "DAS Pago":        brl(row["das_pago"]),
                "DAS Correto":     brl(row["das_correto"]),
                "Credito Mensal":  brl(row["credito"]),
            })

        df_apuracao = pd.DataFrame(linhas)
        st.dataframe(df_apuracao, use_container_width=True, hide_index=True)

        # Totais
        st.markdown("---")
        t1, t2, t3, t4 = st.columns(4)
        t1.metric("Aliquota Efetiva",    pct(ef))
        t2.metric("Total DAS Pago",      brl(sum(r["das_pago"]    for r in apuracao)))
        t3.metric("Total DAS Correto",   brl(sum(r["das_correto"] for r in apuracao)))
        t4.metric("Credito Total Apurado", brl(total_credito),
                  delta="a recuperar" if total_credito > 0 else None)

        if total_credito <= 0:
            st.info("Nenhum credito apurado no periodo. Verifique os XMLs e o RBT12.")

        # Alertas
        sem_data = [i for i in itens if i["mes"] == "SEM-DATA"]
        if sem_data:
            st.warning(
                "{} item(ns) sem data de emissao (dhEmi ausente). "
                "Esses itens foram agrupados como 'SEM-DATA'.".format(len(sem_data))
            )

        # Exportar CSV da apuracao
        csv_bytes = gerar_csv([
            [r["mes"], r["receita_total"], r["receita_monofasica"],
             r["receita_tributavel"], r["das_pago"], r["das_correto"], r["credito"]]
            for r in apuracao
        ])
        st.download_button(
            label="Baixar apuracao CSV",
            data=csv_bytes,
            file_name="apuracao_pgdas.csv",
            mime="text/csv",
        )

    # ── INCONSISTENCIAS ───────────────────────────────────────
    inconsist = df[df["classificacao"] == "INCONSISTENCIA"]
    if not inconsist.empty:
        st.subheader("Inconsistencias Encontradas")
        st.warning("{} item(ns) com NCM ausente ou invalido.".format(len(inconsist)))
        st.dataframe(
            inconsist[["arquivo", "descricao", "ncm", "valor", "motivo"]].rename(columns={
                "arquivo":"Arquivo","descricao":"Descricao",
                "ncm":"NCM","valor":"Valor (R$)","motivo":"Motivo"
            }),
            use_container_width=True,
        )

    # ── TABELA DE ITENS ───────────────────────────────────────
    st.subheader("5  Itens Classificados")
    filtro = st.selectbox(
        "Filtrar por classificacao",
        ["Todos", "MONOFASICO", "NAO MONOFASICO", "INCONSISTENCIA"],
    )
    df_exib = df if filtro == "Todos" else df[df["classificacao"] == filtro]
    st.dataframe(
        df_exib[["mes","arquivo","descricao","ncm","valor","classificacao","motivo"]].rename(columns={
            "mes":"Mes","arquivo":"Arquivo","descricao":"Descricao","ncm":"NCM",
            "valor":"Valor (R$)","classificacao":"Classificacao","motivo":"Motivo"
        }),
        use_container_width=True,
        height=400,
    )

    # ── EXPORTAR EXCEL ────────────────────────────────────────
    st.subheader("6  Exportar Relatorio Completo")
    st.download_button(
        label="Baixar relatorio Excel",
        data=gerar_excel(itens, resumo),
        file_name="relatorio_pis_cofins_monofasico.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.caption(
        "Aviso: MVP para analise preliminar. Valores devem ser validados "
        "por contador ou consultor tributario habilitado. "
        "Calculos baseados na LC 123/2006 e Resolucao CGSN 140/2018."
    )


if __name__ == "__main__":
    main()

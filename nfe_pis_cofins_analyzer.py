"""
=============================================================
  ANALISADOR PIS/COFINS MONOFASICO + APURACAO PGDAS-D v5
  Simples Nacional · Regime de Revenda · CST 04
  RBT12 dinamico · DAS Real vs Estimado · Cruzamento PGDAS
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
from io import BytesIO
from datetime import datetime

# ─────────────────────────────────────────────────────────────
#  1. TABELA DE NCM MONOFASICOS (Tabela 4.3.10 EFD-Contribuicoes)
# ─────────────────────────────────────────────────────────────
TABELA_NCM_MONOFASICO = {
    "27101112": "Gasolina automotiva comum",
    "27101113": "Gasolina automotiva premium",
    "27101121": "Querosene de aviacao",
    "27101131": "Oleo diesel",
    "27101500": "Oleos lubrificantes",
    "27111100": "Gas natural liquefeito (GNL)",
    "27111910": "Gas liquefeito de petroleo (GLP)",
    "30011000": "Glandulas e orgaos para usos opoterapicos",
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
#  2. TABELAS DO SIMPLES NACIONAL – LC 123/2006
# ─────────────────────────────────────────────────────────────
ANEXO_I = [
    {"faixa": 1, "limite": 180_000,   "aliquota": 0.04,  "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,   "aliquota": 0.073, "deducao": 5_940.0},
    {"faixa": 3, "limite": 720_000,   "aliquota": 0.095, "deducao": 13_860.0},
    {"faixa": 4, "limite": 1_800_000, "aliquota": 0.107, "deducao": 22_500.0},
    {"faixa": 5, "limite": 3_600_000, "aliquota": 0.143, "deducao": 87_300.0},
    {"faixa": 6, "limite": 4_800_000, "aliquota": 0.19,  "deducao": 378_000.0},
]
ANEXO_II = [
    {"faixa": 1, "limite": 180_000,   "aliquota": 0.045, "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,   "aliquota": 0.078, "deducao": 5_940.0},
    {"faixa": 3, "limite": 720_000,   "aliquota": 0.10,  "deducao": 13_860.0},
    {"faixa": 4, "limite": 1_800_000, "aliquota": 0.113, "deducao": 22_500.0},
    {"faixa": 5, "limite": 3_600_000, "aliquota": 0.147, "deducao": 85_500.0},
    {"faixa": 6, "limite": 4_800_000, "aliquota": 0.30,  "deducao": 720_000.0},
]
ANEXO_III = [
    {"faixa": 1, "limite": 180_000,   "aliquota": 0.06,  "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,   "aliquota": 0.112, "deducao": 9_360.0},
    {"faixa": 3, "limite": 720_000,   "aliquota": 0.135, "deducao": 17_640.0},
    {"faixa": 4, "limite": 1_800_000, "aliquota": 0.16,  "deducao": 35_640.0},
    {"faixa": 5, "limite": 3_600_000, "aliquota": 0.21,  "deducao": 125_640.0},
    {"faixa": 6, "limite": 4_800_000, "aliquota": 0.33,  "deducao": 648_000.0},
]
ANEXO_IV = [
    {"faixa": 1, "limite": 180_000,   "aliquota": 0.045, "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,   "aliquota": 0.09,  "deducao": 8_100.0},
    {"faixa": 3, "limite": 720_000,   "aliquota": 0.102, "deducao": 12_420.0},
    {"faixa": 4, "limite": 1_800_000, "aliquota": 0.14,  "deducao": 39_780.0},
    {"faixa": 5, "limite": 3_600_000, "aliquota": 0.22,  "deducao": 183_780.0},
    {"faixa": 6, "limite": 4_800_000, "aliquota": 0.33,  "deducao": 828_000.0},
]
ANEXO_V = [
    {"faixa": 1, "limite": 180_000,   "aliquota": 0.15,  "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,   "aliquota": 0.18,  "deducao": 5_400.0},
    {"faixa": 3, "limite": 720_000,   "aliquota": 0.195, "deducao": 13_500.0},
    {"faixa": 4, "limite": 1_800_000, "aliquota": 0.205, "deducao": 20_700.0},
    {"faixa": 5, "limite": 3_600_000, "aliquota": 0.23,  "deducao": 62_100.0},
    {"faixa": 6, "limite": 4_800_000, "aliquota": 0.305, "deducao": 540_000.0},
]
TABELAS_SIMPLES = {
    "Anexo I – Comercio":     ANEXO_I,
    "Anexo II – Industria":   ANEXO_II,
    "Anexo III – Servicos A": ANEXO_III,
    "Anexo IV – Servicos B":  ANEXO_IV,
    "Anexo V – Servicos C":   ANEXO_V,
}

# ─────────────────────────────────────────────────────────────
#  3. TABELA DE REPARTICAO DO DAS – LC 123/2006 / CGSN 140/2018
# ─────────────────────────────────────────────────────────────
REPARTICAO = {
    "Anexo I – Comercio": {
        1: {"pis": 0.0,    "cofins": 0.0},
        2: {"pis": 0.0276, "cofins": 0.1274},
        3: {"pis": 0.0276, "cofins": 0.1274},
        4: {"pis": 0.0276, "cofins": 0.1274},
        5: {"pis": 0.0276, "cofins": 0.1274},
        6: {"pis": 0.0276, "cofins": 0.1274},
    },
    "Anexo II – Industria": {
        1: {"pis": 0.0,    "cofins": 0.0},
        2: {"pis": 0.0186, "cofins": 0.086},
        3: {"pis": 0.0186, "cofins": 0.086},
        4: {"pis": 0.0186, "cofins": 0.086},
        5: {"pis": 0.0186, "cofins": 0.086},
        6: {"pis": 0.0186, "cofins": 0.086},
    },
    "Anexo III – Servicos A": {
        1: {"pis": 0.0,    "cofins": 0.0},
        2: {"pis": 0.0167, "cofins": 0.0773},
        3: {"pis": 0.0167, "cofins": 0.0773},
        4: {"pis": 0.0167, "cofins": 0.0773},
        5: {"pis": 0.0167, "cofins": 0.0773},
        6: {"pis": 0.0167, "cofins": 0.0773},
    },
    "Anexo IV – Servicos B": {
        1: {"pis": 0.0,    "cofins": 0.0},
        2: {"pis": 0.0167, "cofins": 0.0773},
        3: {"pis": 0.0167, "cofins": 0.0773},
        4: {"pis": 0.0167, "cofins": 0.0773},
        5: {"pis": 0.0167, "cofins": 0.0773},
        6: {"pis": 0.0167, "cofins": 0.0773},
    },
    "Anexo V – Servicos C": {
        1: {"pis": 0.0,    "cofins": 0.0},
        2: {"pis": 0.0098, "cofins": 0.0454},
        3: {"pis": 0.0098, "cofins": 0.0454},
        4: {"pis": 0.0098, "cofins": 0.0454},
        5: {"pis": 0.0098, "cofins": 0.0454},
        6: {"pis": 0.0098, "cofins": 0.0454},
    },
}

ALIQUOTA_PIS_COFINS_ESTIMATIVA = 0.0925
DIVERGENCIA_LIMITE = 0.05   # 5% – limiar para alerta de divergencia


# ─────────────────────────────────────────────────────────────
#  4. HELPERS XML
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
    dh = text_local(root, "dhEmi") or text_local(root, "dEmi")
    if not dh:
        return "SEM-DATA"
    try:
        return dh[:7]
    except Exception:
        return "SEM-DATA"

def ler_xml_nfe(conteudo):
    """Le bytes de NF-e. Retorna (lista_itens, mes_ano). Robusto a BOM e namespace."""
    if isinstance(conteudo, bytes):
        conteudo = conteudo.lstrip(b"\xef\xbb\xbf").strip()
    try:
        root = ET.fromstring(conteudo)
    except ET.ParseError as e:
        raise ValueError("XML invalido: {}".format(e))
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
#  5. LEITURA DO PGDAS/DAS  ← NOVO
# ─────────────────────────────────────────────────────────────
COLUNAS_PGDAS_ESPERADAS = ["Mes", "Receita_PGDAS", "DAS_Pago"]

def ler_pgdas(file):
    """
    Le arquivo CSV ou Excel com dados reais do PGDAS/DAS.

    Colunas esperadas (case-insensitive, espaços tolerados):
        Mes           → formato YYYY-MM ou MM/YYYY
        Receita_PGDAS → receita declarada no PGDAS (float)
        DAS_Pago      → valor efetivamente pago do DAS (float)

    Retorna DataFrame padronizado com colunas:
        mes (str YYYY-MM) | receita_pgdas (float) | das_pago_real (float)

    Lanca ValueError com mensagem amigavel em caso de erro.
    """
    try:
        nome = file.name.lower()
        if nome.endswith(".csv"):
            # Tenta separadores comuns: ; e ,
            try:
                df = pd.read_csv(file, sep=";", decimal=",", dtype=str)
                if df.shape[1] < 2:
                    file.seek(0)
                    df = pd.read_csv(file, sep=",", decimal=".", dtype=str)
            except Exception:
                file.seek(0)
                df = pd.read_csv(file, sep=",", decimal=".", dtype=str)
        else:
            df = pd.read_excel(file, dtype=str)
    except Exception as e:
        raise ValueError("Nao foi possivel ler o arquivo: {}".format(e))

    # Normaliza nomes das colunas
    df.columns = [c.strip().replace(" ", "_") for c in df.columns]
    col_map = {c.lower(): c for c in df.columns}

    mapa = {}
    for esperada in COLUNAS_PGDAS_ESPERADAS:
        chave = esperada.lower()
        if chave in col_map:
            mapa[esperada] = col_map[chave]
        else:
            raise ValueError(
                "Coluna '{}' nao encontrada. Colunas no arquivo: {}".format(
                    esperada, list(df.columns)
                )
            )

    df = df.rename(columns={v: k for k, v in mapa.items()})
    df = df[COLUNAS_PGDAS_ESPERADAS].copy()

    # Converte Mes para YYYY-MM
    def normalizar_mes(v):
        v = str(v).strip()
        # Tenta YYYY-MM
        try:
            datetime.strptime(v, "%Y-%m")
            return v
        except ValueError:
            pass
        # Tenta MM/YYYY
        try:
            return datetime.strptime(v, "%m/%Y").strftime("%Y-%m")
        except ValueError:
            pass
        # Tenta MM/AAAA com qualquer separador
        for fmt in ("%m-%Y", "%Y/%m"):
            try:
                return datetime.strptime(v, fmt).strftime("%Y-%m")
            except ValueError:
                continue
        return v  # mantém original se nao reconheceu

    df["Mes"] = df["Mes"].apply(normalizar_mes)

    # Converte valores monetarios
    def parse_brl(v):
        s = str(v).strip().replace("R$", "").replace(" ", "")
        # Detecta formato brasileiro: 1.234,56
        if "," in s and "." in s:
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0

    df["Receita_PGDAS"] = df["Receita_PGDAS"].apply(parse_brl)
    df["DAS_Pago"]      = df["DAS_Pago"].apply(parse_brl)

    df = df.rename(columns={
        "Mes":           "mes",
        "Receita_PGDAS": "receita_pgdas",
        "DAS_Pago":      "das_pago_real",
    })

    df = df.sort_values("mes").reset_index(drop=True)
    return df


# ─────────────────────────────────────────────────────────────
#  6. CRUZAMENTO XML vs PGDAS  ← NOVO
# ─────────────────────────────────────────────────────────────
def cruzar_xml_pgdas(agrupamento_mensal, df_pgdas):
    """
    Cruza os dados de receita apurados dos XMLs com os dados reais do PGDAS.

    Retorna lista de dicts com campos adicionais por mes:
        receita_pgdas   → receita declarada no PGDAS (None se ausente)
        das_pago_real   → DAS efetivamente pago (None se ausente)
        divergencia_pct → |xml - pgdas| / pgdas (None se sem PGDAS)
        tem_pgdas       → bool
        alerta_divergencia → bool (divergencia > DIVERGENCIA_LIMITE)
    """
    pgdas_idx = {}
    if df_pgdas is not None and not df_pgdas.empty:
        for _, row in df_pgdas.iterrows():
            pgdas_idx[str(row["mes"]).strip()] = row

    resultado = []
    for row in agrupamento_mensal:
        mes = row["mes"]
        pgdas_row = pgdas_idx.get(mes)

        if pgdas_row is not None:
            rec_pgdas    = float(pgdas_row["receita_pgdas"])
            das_real     = float(pgdas_row["das_pago_real"])
            rec_xml      = row["receita_total"]
            if rec_pgdas > 0:
                div = abs(rec_xml - rec_pgdas) / rec_pgdas
            else:
                div = 0.0
            tem_pgdas        = True
            alerta_div       = div > DIVERGENCIA_LIMITE
        else:
            rec_pgdas    = None
            das_real     = None
            div          = None
            tem_pgdas    = False
            alerta_div   = False

        resultado.append({
            **row,
            "receita_pgdas":       rec_pgdas,
            "das_pago_real":       das_real,
            "divergencia_pct":     div,
            "tem_pgdas":           tem_pgdas,
            "alerta_divergencia":  alerta_div,
        })
    return resultado


# ─────────────────────────────────────────────────────────────
#  7. CLASSIFICACAO NCM
# ─────────────────────────────────────────────────────────────
def classificar_item(ncm, tabela):
    ncm_limpo = str(ncm).strip()
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
#  8. AGRUPAMENTO MENSAL
# ─────────────────────────────────────────────────────────────
def agrupar_por_mes(itens_classificados):
    meses = {}
    for item in itens_classificados:
        m = item.get("mes", "SEM-DATA")
        if m not in meses:
            meses[m] = {"receita_total": 0.0, "receita_monofasica": 0.0}
        meses[m]["receita_total"] += item["valor"]
        if item["classificacao"] == "MONOFASICO":
            meses[m]["receita_monofasica"] += item["valor"]
    resultado = []
    chaves_validas   = sorted(k for k in meses if k != "SEM-DATA")
    chaves_invalidas = [k for k in meses if k == "SEM-DATA"]
    for m in chaves_validas + chaves_invalidas:
        total = meses[m]["receita_total"]
        mono  = meses[m]["receita_monofasica"]
        resultado.append({
            "mes":                m,
            "receita_total":      total,
            "receita_monofasica": mono,
            "receita_tributavel": total - mono,
        })
    return resultado


# ─────────────────────────────────────────────────────────────
#  9. RBT12 DINAMICO
# ─────────────────────────────────────────────────────────────
def calcular_rbt12(agrupamento_mensal, rbt12_inicial=0.0):
    """
    RBT12 rolling de 12 meses para cada mes.
    rbt12_inicial = proxy dos 12 meses anteriores ao periodo dos XMLs.
    """
    meses    = [r["mes"] for r in agrupamento_mensal]
    receitas = {r["mes"]: r["receita_total"] for r in agrupamento_mensal}
    resultado = {}
    historico = []
    saldo = rbt12_inicial
    for mes in meses:
        resultado[mes] = round(saldo, 2)
        historico.append((mes, receitas[mes]))
        saldo += receitas[mes]
        if len(historico) > 12:
            _, rec_antiga = historico.pop(0)
            saldo -= rec_antiga
    return resultado


# ─────────────────────────────────────────────────────────────
#  10. FAIXA E ALIQUOTA EFETIVA
# ─────────────────────────────────────────────────────────────
def identificar_faixa(rbt12, tabela_anexo):
    if rbt12 <= 0:
        return tabela_anexo[0]
    for faixa in tabela_anexo:
        if rbt12 <= faixa["limite"]:
            return faixa
    return None

def calcular_aliquota_efetiva(rbt12, faixa):
    if rbt12 <= 0:
        return 0.0
    return (rbt12 * faixa["aliquota"] - faixa["deducao"]) / rbt12

def calcular_pct_pis_cofins(faixa_num, nome_anexo):
    rep = REPARTICAO.get(nome_anexo, {})
    f   = rep.get(faixa_num, {})
    return f.get("pis", 0.0) + f.get("cofins", 0.0)


# ─────────────────────────────────────────────────────────────
#  11. PIPELINE DE APURACAO (estimado – mantido do v4)
# ─────────────────────────────────────────────────────────────
def apurar_periodo(agrupamento_mensal, nome_anexo, rbt12_inicial=0.0):
    """Pipeline original v4 – sem dados do PGDAS (modo estimado)."""
    tabela_anexo  = TABELAS_SIMPLES[nome_anexo]
    rbt12_por_mes = calcular_rbt12(agrupamento_mensal, rbt12_inicial)
    resultado = []
    for row in agrupamento_mensal:
        mes   = row["mes"]
        rbt12 = rbt12_por_mes.get(mes, 0.0)
        faixa = identificar_faixa(rbt12, tabela_anexo)
        if faixa is None:
            resultado.append({**row, "rbt12": rbt12, "faixa": "ACIMA",
                               "aliquota_efetiva": None, "das_pago": None,
                               "das_correto": None, "credito_bruto": None,
                               "pct_pis_cofins": None, "credito_real": None,
                               "alerta": "RBT12 acima de R$ 4,8M"})
            continue
        aliq_ef     = calcular_aliquota_efetiva(rbt12, faixa)
        das_pago    = row["receita_total"]     * aliq_ef
        das_correto = row["receita_tributavel"] * aliq_ef
        cred_bruto  = das_pago - das_correto
        pct_pc      = calcular_pct_pis_cofins(faixa["faixa"], nome_anexo)
        resultado.append({
            **row,
            "rbt12":            rbt12,
            "faixa":            faixa["faixa"],
            "aliquota_nominal": faixa["aliquota"],
            "aliquota_efetiva": aliq_ef,
            "das_pago":         das_pago,
            "das_correto":      das_correto,
            "credito_bruto":    cred_bruto,
            "pct_pis_cofins":   pct_pc,
            "credito_real":     cred_bruto * pct_pc,
            "alerta":           "",
        })
    return resultado


# ─────────────────────────────────────────────────────────────
#  12. PIPELINE REAL (com PGDAS)  ← NOVO
# ─────────────────────────────────────────────────────────────
def apurar_periodo_real(agrupamento_mensal, nome_anexo,
                        rbt12_inicial=0.0, df_pgdas=None):
    """
    Evolucao do apurar_periodo() com suporte a DAS real do PGDAS.

    Logica de fonte do DAS por mes:
      - Se o mes tem dados PGDAS → usa das_pago_real (fonte = 'REAL')
      - Caso contrario           → usa das_pago estimado (fonte = 'ESTIMADO')

    Campos adicionais vs. apurar_periodo():
        receita_pgdas      → receita declarada no PGDAS
        das_pago_real      → DAS efetivamente pago
        divergencia_pct    → divergencia receita XML vs PGDAS
        alerta_divergencia → bool
        tem_pgdas          → bool
        fonte_das          → 'REAL' | 'ESTIMADO'
        das_usado          → valor do DAS efetivamente usado no calculo
    """
    tabela_anexo  = TABELAS_SIMPLES[nome_anexo]
    rbt12_por_mes = calcular_rbt12(agrupamento_mensal, rbt12_inicial)

    # Cruzamento com PGDAS
    agrupado_cruzado = cruzar_xml_pgdas(agrupamento_mensal, df_pgdas)

    resultado = []
    for row in agrupado_cruzado:
        mes   = row["mes"]
        rbt12 = rbt12_por_mes.get(mes, 0.0)
        faixa = identificar_faixa(rbt12, tabela_anexo)

        if faixa is None:
            resultado.append({
                **row,
                "rbt12": rbt12, "faixa": "ACIMA",
                "aliquota_nominal": None, "aliquota_efetiva": None,
                "das_estimado": None, "das_usado": None, "fonte_das": "–",
                "das_correto": None, "credito_bruto": None,
                "pct_pis_cofins": None, "credito_real": None,
                "alerta": "RBT12 acima de R$ 4,8M – fora do Simples",
            })
            continue

        aliq_ef      = calcular_aliquota_efetiva(rbt12, faixa)
        das_estimado = row["receita_total"] * aliq_ef

        # Seleciona fonte do DAS
        if row["tem_pgdas"] and row["das_pago_real"] is not None:
            das_usado  = row["das_pago_real"]
            fonte_das  = "REAL"
        else:
            das_usado  = das_estimado
            fonte_das  = "ESTIMADO"

        # DAS correto sempre sobre receita tributavel
        das_correto = row["receita_tributavel"] * aliq_ef
        cred_bruto  = das_usado - das_correto
        pct_pc      = calcular_pct_pis_cofins(faixa["faixa"], nome_anexo)
        cred_real   = cred_bruto * pct_pc

        resultado.append({
            **row,
            "rbt12":            rbt12,
            "faixa":            faixa["faixa"],
            "aliquota_nominal": faixa["aliquota"],
            "aliquota_efetiva": aliq_ef,
            "das_estimado":     das_estimado,
            "das_usado":        das_usado,
            "fonte_das":        fonte_das,
            "das_correto":      das_correto,
            "credito_bruto":    cred_bruto,
            "pct_pis_cofins":   pct_pc,
            "credito_real":     cred_real,
            "alerta":           "",
        })
    return resultado


# ─────────────────────────────────────────────────────────────
#  13. RESUMO GERAL
# ─────────────────────────────────────────────────────────────
def calcular_resumo(itens_classificados, aliquota=ALIQUOTA_PIS_COFINS_ESTIMATIVA):
    total_geral     = sum(i["valor"] for i in itens_classificados)
    total_mono      = sum(i["valor"] for i in itens_classificados if i["classificacao"] == "MONOFASICO")
    total_nao_mono  = sum(i["valor"] for i in itens_classificados if i["classificacao"] == "NAO MONOFASICO")
    total_inconsist = sum(i["valor"] for i in itens_classificados if i["classificacao"] == "INCONSISTENCIA")
    return {
        "total_geral":      total_geral,
        "total_monofasico": total_mono,
        "total_nao_mono":   total_nao_mono,
        "total_inconsist":  total_inconsist,
        "estimativa_recup": total_mono * aliquota,
    }

def processar_xmls(arquivos, tabela, aliquota=ALIQUOTA_PIS_COFINS_ESTIMATIVA):
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
                "arquivo": nome, "mes": mes_ano,
                "descricao": item["descricao"], "ncm": item["ncm_raw"],
                "valor": item["valor"], "classificacao": classif, "motivo": motivo,
            })
    return todos_itens, calcular_resumo(todos_itens, aliquota)


# ─────────────────────────────────────────────────────────────
#  14. EXPORTACOES
# ─────────────────────────────────────────────────────────────
def gerar_excel(itens, resumo):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_i = pd.DataFrame(itens)
        df_i.columns = ["Arquivo","Mes","Descricao","NCM","Valor (R$)","Classificacao","Motivo"]
        df_i.to_excel(writer, sheet_name="Itens", index=False)
        pd.DataFrame([{
            "Fat. Total":         resumo["total_geral"],
            "Fat. Monofasico":    resumo["total_monofasico"],
            "Fat. Nao Monof.":    resumo["total_nao_mono"],
            "Inconsistencias":    resumo["total_inconsist"],
            "Estim. Recuperacao": resumo["estimativa_recup"],
        }]).to_excel(writer, sheet_name="Resumo", index=False)
    return output.getvalue()

def gerar_csv_apuracao(apuracao):
    linhas = []
    for r in apuracao:
        linhas.append({
            "Mes":           r["mes"],
            "RBT12":         r.get("rbt12",""),
            "Faixa":         r.get("faixa",""),
            "Aliq. Efetiva": r.get("aliquota_efetiva",""),
            "Rec. XML":      r.get("receita_total",""),
            "Rec. PGDAS":    r.get("receita_pgdas","") if r.get("receita_pgdas") is not None else "",
            "Divergencia %": "{:.2f}%".format(r["divergencia_pct"]*100) if r.get("divergencia_pct") is not None else "",
            "Rec. Monof.":   r.get("receita_monofasica",""),
            "Rec. Trib.":    r.get("receita_tributavel",""),
            "DAS Usado":     r.get("das_usado",""),
            "Fonte DAS":     r.get("fonte_das",""),
            "DAS Correto":   r.get("das_correto",""),
            "Cred. Bruto":   r.get("credito_bruto",""),
            "% PIS+COF":     r.get("pct_pis_cofins",""),
            "Cred. Real":    r.get("credito_real",""),
        })
    return pd.DataFrame(linhas).to_csv(index=False, sep=";", decimal=",").encode("utf-8")


# ─────────────────────────────────────────────────────────────
#  15. FORMATACAO
# ─────────────────────────────────────────────────────────────
def brl(v):
    if v is None:
        return "–"
    return "R$ {:,.2f}".format(v).replace(",","X").replace(".",",").replace("X",".")

def pct(v, casas=4):
    if v is None:
        return "–"
    return "{:.{}f}%".format(v * 100, casas)

def fmt_mes(m):
    try:
        return datetime.strptime(m, "%Y-%m").strftime("%b/%Y")
    except Exception:
        return m

def badge_fonte(f):
    return "🟢 REAL" if f == "REAL" else "🟡 ESTIMADO"

def badge_div(div):
    if div is None:
        return "–"
    s = "{:.1f}%".format(div * 100)
    return "🔴 " + s if div > DIVERGENCIA_LIMITE else "✅ " + s


# ─────────────────────────────────────────────────────────────
#  16. INTERFACE STREAMLIT
# ─────────────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="PIS/COFINS + PGDAS-D", page_icon="📊", layout="wide")

    st.title("📊 Analisador PIS/COFINS Monofasico + Apuracao PGDAS-D")
    st.caption("Simples Nacional · Regime de Revenda · DAS Real vs Estimado · Cruzamento PGDAS")
    st.divider()

    # ── SIDEBAR ───────────────────────────────────────────────
    with st.sidebar:
        st.header("Configuracoes")

        st.subheader("Simples Nacional")
        nome_anexo = st.selectbox("Anexo", list(TABELAS_SIMPLES.keys()))
        rbt12_inicial = st.number_input(
            "RBT12 anterior ao periodo (R$)",
            min_value=0.0, value=360_000.0, step=1_000.0, format="%.2f",
            help="Receita bruta dos 12 meses anteriores ao primeiro XML."
        )

        st.markdown("---")
        st.subheader("Modo de Calculo")
        usar_pgdas_real = st.checkbox(
            "Usar dados reais do PGDAS",
            value=False,
            help="Quando ativado, o DAS real substitui o DAS estimado no calculo do credito."
        )

        st.markdown("---")
        st.subheader("PIS/COFINS (estimativa)")
        aliquota_pc = st.number_input(
            "Aliquota estimativa (%)", min_value=0.0, max_value=100.0,
            value=9.25, step=0.05, format="%.2f",
        )
        aliquota_decimal = aliquota_pc / 100

        st.markdown("---")
        st.markdown("`{}` NCMs monofasicos na tabela".format(len(TABELA_NCM_MONOFASICO)))

        with st.expander("Reparticao do anexo"):
            rep   = REPARTICAO.get(nome_anexo, {})
            rows_ = [{"Faixa": fn, "PIS": pct(v.get("pis",0), 2),
                      "COFINS": pct(v.get("cofins",0), 2),
                      "PIS+COFINS": pct(v.get("pis",0)+v.get("cofins",0), 2)}
                     for fn, v in rep.items()]
            st.dataframe(pd.DataFrame(rows_), hide_index=True, use_container_width=True)

    # ── UPLOAD XMLs ───────────────────────────────────────────
    st.subheader("1  Upload dos XMLs de NF-e")
    uploaded_xmls = st.file_uploader(
        "Selecione um ou mais arquivos XML",
        type=["xml"], accept_multiple_files=True,
    )
    if not uploaded_xmls:
        st.info("Aguardando upload dos XMLs de NF-e...")
        st.stop()

    # ── UPLOAD PGDAS ──────────────────────────────────────────
    df_pgdas = None
    if usar_pgdas_real:
        st.subheader("2  Upload dos Dados do PGDAS/DAS")

        with st.expander("Formato esperado do arquivo", expanded=False):
            st.markdown("""
O arquivo (CSV ou Excel) deve ter **exatamente estas colunas** (nomes flexiveis, sem acento):

| Mes | Receita_PGDAS | DAS_Pago |
|-----|---------------|----------|
| 2024-01 | 45000,00 | 1980,00 |
| 2024-02 | 52000,00 | 2288,00 |

- **Mes**: formato `YYYY-MM` ou `MM/YYYY`
- **Receita_PGDAS**: receita bruta declarada no PGDAS-D
- **DAS_Pago**: valor pago no DAS (guia gerada pelo PGDAS)
- Separador CSV aceito: `;` ou `,`
- Valores monetarios: virgula ou ponto como decimal
            """)
            # Botao de download do template
            template_csv = "Mes;Receita_PGDAS;DAS_Pago\n2024-01;45000,00;1980,00\n2024-02;52000,00;2288,00\n"
            st.download_button(
                "Baixar template CSV",
                data=template_csv.encode("utf-8"),
                file_name="template_pgdas.csv",
                mime="text/csv",
            )

        uploaded_pgdas = st.file_uploader(
            "Selecione o arquivo CSV ou Excel do PGDAS",
            type=["csv", "xlsx", "xls"],
            key="pgdas_upload",
        )

        if uploaded_pgdas:
            try:
                df_pgdas = ler_pgdas(uploaded_pgdas)
                st.success(
                    "PGDAS carregado: {} mes(es) encontrado(s). "
                    "Periodo: {} a {}".format(
                        len(df_pgdas),
                        fmt_mes(df_pgdas["mes"].min()),
                        fmt_mes(df_pgdas["mes"].max()),
                    )
                )
                st.dataframe(
                    df_pgdas.rename(columns={
                        "mes": "Mes", "receita_pgdas": "Receita PGDAS (R$)",
                        "das_pago_real": "DAS Pago Real (R$)"
                    }),
                    hide_index=True, use_container_width=True,
                )
            except ValueError as e:
                st.error("Erro no arquivo PGDAS: {}".format(e))
                df_pgdas = None
        else:
            st.warning(
                "Arquivo PGDAS nao carregado. "
                "O calculo usara DAS estimado como fallback para todos os meses."
            )

    # ── PROCESSAMENTO ─────────────────────────────────────────
    secao_base = 3 if usar_pgdas_real else 2
    arquivos = [(f.name, f.read()) for f in uploaded_xmls]

    with st.spinner("Processando XMLs..."):
        itens, resumo = processar_xmls(arquivos, TABELA_NCM_MONOFASICO, aliquota_decimal)

    if not itens:
        st.error("Nenhum item extraido dos XMLs enviados.")
        st.stop()

    df_itens = pd.DataFrame(itens)

    # ── METRICAS GERAIS ───────────────────────────────────────
    st.subheader("{}  Resumo Geral".format(secao_base))
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Faturamento Total",          brl(resumo["total_geral"]))
    c2.metric("Faturamento Monofasico",     brl(resumo["total_monofasico"]))
    c3.metric("Faturamento Nao Monofasico", brl(resumo["total_nao_mono"]))
    c4.metric("Estimativa PIS+COFINS ({:.2f}%)".format(aliquota_pc),
              brl(resumo["estimativa_recup"]))

    # ── GRAFICO ───────────────────────────────────────────────
    st.subheader("{}  Composicao do Faturamento".format(secao_base + 1))
    graf = {"Monofasico": resumo["total_monofasico"], "Nao Monofasico": resumo["total_nao_mono"]}
    if resumo["total_inconsist"] > 0:
        graf["Inconsistencias"] = resumo["total_inconsist"]
    st.bar_chart(pd.DataFrame.from_dict(graf, orient="index", columns=["Valor (R$)"]))

    # ── APURACAO PGDAS-D ──────────────────────────────────────
    st.subheader("{}  Apuracao PGDAS-D – Credito Real PIS/COFINS".format(secao_base + 2))

    agrupado = agrupar_por_mes(itens)
    apuracao = apurar_periodo_real(agrupado, nome_anexo, rbt12_inicial, df_pgdas)

    # Alertas globais
    for r in apuracao:
        if r.get("alerta"):
            st.error("Mes {}: {}".format(fmt_mes(r["mes"]), r["alerta"]))
        if r.get("alerta_divergencia"):
            st.warning(
                "Divergencia em {}: Receita XML {} vs PGDAS {} "
                "(divergencia: {})".format(
                    fmt_mes(r["mes"]),
                    brl(r["receita_total"]),
                    brl(r["receita_pgdas"]),
                    badge_div(r["divergencia_pct"]),
                )
            )
        if not r.get("tem_pgdas") and usar_pgdas_real:
            st.info("Mes {} sem dados PGDAS – usando DAS estimado.".format(fmt_mes(r["mes"])))

    # Monta tabela
    linhas = []
    for r in apuracao:
        linha = {
            "Mes":            fmt_mes(r["mes"]),
            "RBT12":          brl(r.get("rbt12")),
            "Faixa":          r.get("faixa", "–"),
            "Aliq. Ef.":      pct(r.get("aliquota_efetiva"), 4),
            "Rec. XML":       brl(r.get("receita_total")),
        }
        if usar_pgdas_real:
            linha["Rec. PGDAS"]   = brl(r.get("receita_pgdas"))
            linha["Divergencia"]  = badge_div(r.get("divergencia_pct"))
        linha.update({
            "Rec. Monof.":    brl(r.get("receita_monofasica")),
            "Rec. Trib.":     brl(r.get("receita_tributavel")),
            "DAS Usado":      brl(r.get("das_usado")),
        })
        if usar_pgdas_real:
            linha["Fonte DAS"] = badge_fonte(r.get("fonte_das", "ESTIMADO"))
        linha.update({
            "DAS Correto":    brl(r.get("das_correto")),
            "Cred. Bruto":    brl(r.get("credito_bruto")),
            "% PIS+COF":      pct(r.get("pct_pis_cofins"), 2),
            "Cred. Real":     brl(r.get("credito_real")),
        })
        linhas.append(linha)

    st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)

    # Totais
    st.markdown("---")
    valid = [r for r in apuracao if r.get("credito_real") is not None]
    tot_das_usado  = sum(r["das_usado"]     for r in valid)
    tot_das_corr   = sum(r["das_correto"]   for r in valid)
    tot_cred_bruto = sum(r["credito_bruto"] for r in valid)
    tot_cred_real  = sum(r["credito_real"]  for r in valid)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total DAS Usado",         brl(tot_das_usado))
    m2.metric("Total DAS Correto",       brl(tot_das_corr))
    m3.metric("Credito Bruto Total",     brl(tot_cred_bruto))
    m4.metric("Credito Real PIS/COFINS", brl(tot_cred_real),
              delta="a recuperar" if tot_cred_real > 0 else None)

    # Legenda modo
    if usar_pgdas_real:
        n_real = sum(1 for r in valid if r.get("fonte_das") == "REAL")
        n_est  = sum(1 for r in valid if r.get("fonte_das") == "ESTIMADO")
        st.caption(
            "🟢 {} mes(es) com DAS real do PGDAS  |  "
            "🟡 {} mes(es) com DAS estimado (fallback)".format(n_real, n_est)
        )

    with st.expander("Metodologia de calculo"):
        st.markdown("""
**RBT12 dinamico:** janela deslizante de 12 meses partindo do RBT12 anterior informado.

**Aliquota efetiva (PGDAS-D):**
```
aliquota_efetiva = (RBT12 x aliquota_nominal - deducao) / RBT12
```
**Calculo do credito:**
```
DAS Correto   = Receita Tributavel x aliquota_efetiva
Credito Bruto = DAS Usado - DAS Correto
Credito Real  = Credito Bruto x (% PIS + % COFINS da reparticao)
```
**Fonte do DAS:** quando os dados reais do PGDAS estao carregados,
o DAS real substitui o estimado. Meses sem PGDAS usam DAS estimado automaticamente.

**Divergencia:** |Rec. XML - Rec. PGDAS| / Rec. PGDAS. Alerta quando > 5%.
        """)

    # Export CSV
    st.download_button(
        "Baixar apuracao CSV",
        data=gerar_csv_apuracao(apuracao),
        file_name="apuracao_pgdas_v5.csv",
        mime="text/csv",
    )

    # ── ALERTAS NCM ───────────────────────────────────────────
    inconsist = df_itens[df_itens["classificacao"] == "INCONSISTENCIA"]
    if not inconsist.empty:
        st.subheader("Inconsistencias de NCM")
        st.warning("{} item(ns) com NCM ausente ou invalido.".format(len(inconsist)))
        st.dataframe(
            inconsist[["arquivo","descricao","ncm","valor","motivo"]].rename(columns={
                "arquivo":"Arquivo","descricao":"Descricao",
                "ncm":"NCM","valor":"Valor (R$)","motivo":"Motivo"
            }),
            use_container_width=True,
        )

    sem_data = df_itens[df_itens["mes"] == "SEM-DATA"]
    if not sem_data.empty:
        st.warning("{} item(ns) sem data de emissao (dhEmi ausente).".format(len(sem_data)))

    # ── TABELA DE ITENS ───────────────────────────────────────
    st.subheader("{}  Itens Classificados".format(secao_base + 3))
    filtro = st.selectbox("Filtrar", ["Todos","MONOFASICO","NAO MONOFASICO","INCONSISTENCIA"])
    df_exib = df_itens if filtro == "Todos" else df_itens[df_itens["classificacao"] == filtro]
    st.dataframe(
        df_exib[["mes","arquivo","descricao","ncm","valor","classificacao","motivo"]].rename(columns={
            "mes":"Mes","arquivo":"Arquivo","descricao":"Descricao","ncm":"NCM",
            "valor":"Valor (R$)","classificacao":"Classificacao","motivo":"Motivo"
        }),
        use_container_width=True, height=400,
    )

    # ── EXPORTAR EXCEL ────────────────────────────────────────
    st.subheader("{}  Exportar Relatorio Completo".format(secao_base + 4))
    st.download_button(
        "Baixar relatorio Excel",
        data=gerar_excel(itens, resumo),
        file_name="relatorio_pis_cofins_v5.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.caption(
        "Aviso: MVP para analise preliminar. Valores devem ser validados "
        "por contador ou consultor tributario habilitado. "
        "LC 123/2006 · Resolucao CGSN 140/2018."
    )

if __name__ == "__main__":
    main()

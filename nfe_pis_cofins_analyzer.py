"""
=============================================================
  PIS/COFINS PRO – Apuracao PGDAS-D  |  v6
  Simples Nacional · Regime de Revenda · CST 04
  Validacao Tributaria · Score de Oportunidade · Auditoria
=============================================================
Dependencias:
    pip install streamlit pandas openpyxl plotly

Execucao:
    streamlit run app.py

Arquitetura preparada para expansao:
    services/   → logica de negocio (xml, pgdas, calculo)
    validators/ → validacao tributaria
    exports/    → csv, excel, relatorios
    utils/      → formatacao, logs, helpers
=============================================================
"""

from __future__ import annotations
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime
from typing import Optional

# ─────────────────────────────────────────────────────────────
#  CONSTANTES GLOBAIS
# ─────────────────────────────────────────────────────────────
VERSAO          = "6.0"
DIVERGENCIA_MAX = 0.05      # 5% → alerta de divergencia
UPLOAD_MAX_MB   = 50        # limite por arquivo XML
ALIQ_ESTIMATIVA = 0.0925    # 9,25% estimativa PIS+COFINS

# CSTs de PIS/COFINS compatíveis com monofasico (revenda)
CST_MONOFASICO_VALIDOS    = {"04", "4"}
# CSTs que sinalizam tributacao normal (incompatível com monofasico)
CST_TRIBUTACAO_NORMAL     = {"01", "1", "02", "2", "05", "5", "06", "6",
                              "07", "7", "08", "8", "09", "9"}
# CFOPs de revenda (compra para revenda no mercado interno)
CFOP_REVENDA_VALIDOS      = {"5102","5403","5405","6102","6403","6404",
                              "5101","6101","5104","6104"}
# CFOPs de devolucao, remessa, etc. – nao devem ter monofasico
CFOP_INCOMPATIVEL_MONO    = {"5201","5202","6201","6202","5910","6910",
                              "5949","6949"}

# ─────────────────────────────────────────────────────────────
#  TABELA NCM MONOFASICOS (Tabela 4.3.10 EFD-Contribuicoes)
# ─────────────────────────────────────────────────────────────
TABELA_NCM_MONOFASICO: dict[str, str] = {
    "27101112": "Gasolina automotiva comum",
    "27101113": "Gasolina automotiva premium",
    "27101121": "Querosene de aviacao",
    "27101131": "Oleo diesel",
    "27101500": "Oleos lubrificantes",
    "27111100": "Gas natural liquefeito (GNL)",
    "27111910": "Gas liquefeito de petroleo (GLP)",
    "30011000": "Glandulas/orgaos opoterapicos",
    "30021000": "Antissoros e imunoglobulinas",
    "30022000": "Vacinas medicina humana",
    "30023000": "Vacinas medicina veterinaria",
    "30031000": "Medicamentos c/ penicilinas",
    "30032000": "Medicamentos c/ antibioticos",
    "30039099": "Outros medicamentos mistura",
    "30041000": "Medicamentos penicilinas (doses)",
    "30042000": "Medicamentos antibioticos (doses)",
    "30043900": "Medicamentos hormonais",
    "30049099": "Outros medicamentos uso humano",
    "33011000": "Oleos essenciais citricos",
    "33012900": "Outros oleos essenciais",
    "33030010": "Perfumes (extratos)",
    "33030020": "Aguas-de-colonia",
    "33041000": "Maquiagem para labios",
    "33042000": "Sombras e delineadores",
    "33049900": "Outros produtos de beleza",
    "33051000": "Xampus",
    "33052000": "Preparacoes ondulacao/alisamento",
    "33053000": "Laques",
    "33059000": "Outras preparacoes capilares",
    "33061000": "Dentifricio",
    "33062000": "Fio dental",
    "33069000": "Higiene bucal outros",
    "33071000": "Preparacoes para barbear",
    "33072000": "Desodorantes e antiperspirantes",
    "33074900": "Outros produtos toucador",
    "22011000": "Agua mineral/gaseificada",
    "22019000": "Outras aguas",
    "22021000": "Agua c/ acucar/adocante",
    "22029000": "Bebidas nao alcoolicas",
    "22030000": "Cerveja de malte",
    "22060000": "Bebidas fermentadas",
    "22071000": "Alcool etilico >= 80%",
    "22082000": "Conhaque",
    "22083000": "Uisque",
    "22084000": "Rum e tafia",
    "22085000": "Gim e genebra",
    "22086000": "Vodca",
    "22087000": "Licores",
    "22089900": "Outras bebidas alcoolicas",
    "87031000": "Veiculos neve/quadriciclos",
    "87032100": "Automoveis <= 1000 cm3",
    "87032200": "Automoveis 1000-1500 cm3",
    "87032300": "Automoveis 1500-3000 cm3",
    "87032400": "Automoveis > 3000 cm3",
    "87033300": "Automoveis diesel > 2500 cm3",
    "87060010": "Chassis c/ motor",
    "87089900": "Acessorios veiculos",
    "87111000": "Motos <= 50 cm3",
    "87112000": "Motos 50-250 cm3",
    "87113000": "Motos 250-500 cm3",
    "87114000": "Motos 500-800 cm3",
    "87115000": "Motos > 800 cm3",
    "40111000": "Pneus novos automoveis",
    "40112000": "Pneus novos onibus/caminhoes",
    "40113000": "Pneus novos avioes",
    "40114000": "Pneus novos motocicletas",
    "40119100": "Pneus novos outros",
    "40121100": "Pneus recauchutados automoveis",
    "40121200": "Pneus recauchutados onibus/caminhoes",
}

# ─────────────────────────────────────────────────────────────
#  TABELAS SIMPLES NACIONAL – LC 123/2006
# ─────────────────────────────────────────────────────────────
ANEXO_I = [
    {"faixa":1,"limite":180_000,   "aliquota":0.04,  "deducao":0.0},
    {"faixa":2,"limite":360_000,   "aliquota":0.073, "deducao":5_940.0},
    {"faixa":3,"limite":720_000,   "aliquota":0.095, "deducao":13_860.0},
    {"faixa":4,"limite":1_800_000, "aliquota":0.107, "deducao":22_500.0},
    {"faixa":5,"limite":3_600_000, "aliquota":0.143, "deducao":87_300.0},
    {"faixa":6,"limite":4_800_000, "aliquota":0.19,  "deducao":378_000.0},
]
ANEXO_II = [
    {"faixa":1,"limite":180_000,   "aliquota":0.045, "deducao":0.0},
    {"faixa":2,"limite":360_000,   "aliquota":0.078, "deducao":5_940.0},
    {"faixa":3,"limite":720_000,   "aliquota":0.10,  "deducao":13_860.0},
    {"faixa":4,"limite":1_800_000, "aliquota":0.113, "deducao":22_500.0},
    {"faixa":5,"limite":3_600_000, "aliquota":0.147, "deducao":85_500.0},
    {"faixa":6,"limite":4_800_000, "aliquota":0.30,  "deducao":720_000.0},
]
ANEXO_III = [
    {"faixa":1,"limite":180_000,   "aliquota":0.06,  "deducao":0.0},
    {"faixa":2,"limite":360_000,   "aliquota":0.112, "deducao":9_360.0},
    {"faixa":3,"limite":720_000,   "aliquota":0.135, "deducao":17_640.0},
    {"faixa":4,"limite":1_800_000, "aliquota":0.16,  "deducao":35_640.0},
    {"faixa":5,"limite":3_600_000, "aliquota":0.21,  "deducao":125_640.0},
    {"faixa":6,"limite":4_800_000, "aliquota":0.33,  "deducao":648_000.0},
]
ANEXO_IV = [
    {"faixa":1,"limite":180_000,   "aliquota":0.045, "deducao":0.0},
    {"faixa":2,"limite":360_000,   "aliquota":0.09,  "deducao":8_100.0},
    {"faixa":3,"limite":720_000,   "aliquota":0.102, "deducao":12_420.0},
    {"faixa":4,"limite":1_800_000, "aliquota":0.14,  "deducao":39_780.0},
    {"faixa":5,"limite":3_600_000, "aliquota":0.22,  "deducao":183_780.0},
    {"faixa":6,"limite":4_800_000, "aliquota":0.33,  "deducao":828_000.0},
]
ANEXO_V = [
    {"faixa":1,"limite":180_000,   "aliquota":0.15,  "deducao":0.0},
    {"faixa":2,"limite":360_000,   "aliquota":0.18,  "deducao":5_400.0},
    {"faixa":3,"limite":720_000,   "aliquota":0.195, "deducao":13_500.0},
    {"faixa":4,"limite":1_800_000, "aliquota":0.205, "deducao":20_700.0},
    {"faixa":5,"limite":3_600_000, "aliquota":0.23,  "deducao":62_100.0},
    {"faixa":6,"limite":4_800_000, "aliquota":0.305, "deducao":540_000.0},
]
TABELAS_SIMPLES: dict[str, list] = {
    "Anexo I – Comercio":     ANEXO_I,
    "Anexo II – Industria":   ANEXO_II,
    "Anexo III – Servicos A": ANEXO_III,
    "Anexo IV – Servicos B":  ANEXO_IV,
    "Anexo V – Servicos C":   ANEXO_V,
}

REPARTICAO: dict[str, dict[int, dict]] = {
    "Anexo I – Comercio":     {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0276,"cofins":0.1274},3:{"pis":0.0276,"cofins":0.1274},4:{"pis":0.0276,"cofins":0.1274},5:{"pis":0.0276,"cofins":0.1274},6:{"pis":0.0276,"cofins":0.1274}},
    "Anexo II – Industria":   {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0186,"cofins":0.086}, 3:{"pis":0.0186,"cofins":0.086}, 4:{"pis":0.0186,"cofins":0.086}, 5:{"pis":0.0186,"cofins":0.086}, 6:{"pis":0.0186,"cofins":0.086}},
    "Anexo III – Servicos A": {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0167,"cofins":0.0773},3:{"pis":0.0167,"cofins":0.0773},4:{"pis":0.0167,"cofins":0.0773},5:{"pis":0.0167,"cofins":0.0773},6:{"pis":0.0167,"cofins":0.0773}},
    "Anexo IV – Servicos B":  {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0167,"cofins":0.0773},3:{"pis":0.0167,"cofins":0.0773},4:{"pis":0.0167,"cofins":0.0773},5:{"pis":0.0167,"cofins":0.0773},6:{"pis":0.0167,"cofins":0.0773}},
    "Anexo V – Servicos C":   {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0098,"cofins":0.0454},3:{"pis":0.0098,"cofins":0.0454},4:{"pis":0.0098,"cofins":0.0454},5:{"pis":0.0098,"cofins":0.0454},6:{"pis":0.0098,"cofins":0.0454}},
}

# ─────────────────────────────────────────────────────────────
#  UTILS – formatacao
# ─────────────────────────────────────────────────────────────
def brl(v: Optional[float]) -> str:
    if v is None: return "–"
    return "R$ {:,.2f}".format(v).replace(",","X").replace(".",",").replace("X",".")

def pct_fmt(v: Optional[float], casas: int = 2) -> str:
    if v is None: return "–"
    return "{:.{}f}%".format(v * 100, casas)

def fmt_mes(m: str) -> str:
    try: return datetime.strptime(m, "%Y-%m").strftime("%b/%Y")
    except: return m

def r2(v: Optional[float]) -> Optional[float]:
    """Arredondamento contabil padrao."""
    return round(v, 2) if v is not None else None


# ─────────────────────────────────────────────────────────────
#  UTILS – logs de auditoria
# ─────────────────────────────────────────────────────────────
_LOGS: list[dict] = []

def log(nivel: str, categoria: str, mensagem: str, detalhe: str = "") -> None:
    """Registra evento de auditoria."""
    _LOGS.append({
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "nivel":     nivel,       # INFO | AVISO | ERRO | RISCO
        "categoria": categoria,
        "mensagem":  mensagem,
        "detalhe":   detalhe,
    })

def get_logs() -> list[dict]:
    return list(_LOGS)

def limpar_logs() -> None:
    _LOGS.clear()


# ─────────────────────────────────────────────────────────────
#  VALIDATORS – validacao tributaria avancada
# ─────────────────────────────────────────────────────────────
def validar_item_tributario(
    ncm: str,
    cst_pis: str,
    cst_cofins: str,
    cfop: str,
    classif_ncm: str,
) -> dict:
    """
    Valida o item considerando NCM + CST PIS/COFINS + CFOP.

    Retorna:
        status_tributario : MONOFASICO_VALIDADO | MONOFASICO_COM_RISCO |
                            NAO_MONOFASICO | INCONSISTENTE
        score_risco       : 0 (sem risco) a 100 (alto risco)
        motivo_alerta     : explicacao textual
    """
    alertas: list[str] = []
    score = 0

    is_mono_ncm  = (classif_ncm == "MONOFASICO")
    cfop_limpo   = cfop.strip().replace(".", "")
    cst_p        = cst_pis.strip().lstrip("0") or "0"
    cst_c        = cst_cofins.strip().lstrip("0") or "0"

    # ── Caso 1: NCM nao monofasico – verificacao simples ──────
    if not is_mono_ncm:
        # Alerta se CST marcado como monofasico mas NCM nao consta
        if cst_p in CST_MONOFASICO_VALIDOS or cst_c in CST_MONOFASICO_VALIDOS:
            alertas.append("CST PIS/COFINS indica monofasico mas NCM nao consta na tabela")
            score += 40
            log("RISCO", "ValidacaoTributaria",
                "CST monofasico com NCM nao monofasico", "NCM={}".format(ncm))
            return {"status_tributario": "INCONSISTENTE",
                    "score_risco": score,
                    "motivo_alerta": " | ".join(alertas)}
        return {"status_tributario": "NAO_MONOFASICO", "score_risco": 0, "motivo_alerta": ""}

    # ── Caso 2: NCM monofasico – validacao cruzada ─────────────
    # 2a. Verifica CST PIS
    if cst_p and cst_p not in CST_MONOFASICO_VALIDOS and cst_p != "0":
        if cst_p in CST_TRIBUTACAO_NORMAL:
            alertas.append("CST PIS {} incompativel com regime monofasico".format(cst_pis))
            score += 35
        else:
            alertas.append("CST PIS {} nao reconhecido".format(cst_pis))
            score += 15

    # 2b. Verifica CST COFINS
    if cst_c and cst_c not in CST_MONOFASICO_VALIDOS and cst_c != "0":
        if cst_c in CST_TRIBUTACAO_NORMAL:
            alertas.append("CST COFINS {} incompativel com regime monofasico".format(cst_cofins))
            score += 35
        else:
            alertas.append("CST COFINS {} nao reconhecido".format(cst_cofins))
            score += 15

    # 2c. Verifica CFOP
    if cfop_limpo:
        if cfop_limpo in CFOP_INCOMPATIVEL_MONO:
            alertas.append("CFOP {} incompativel com revenda monofasica".format(cfop))
            score += 30
            log("RISCO", "ValidacaoTributaria",
                "CFOP incompativel com monofasico", "NCM={} CFOP={}".format(ncm, cfop))
        elif cfop_limpo and cfop_limpo not in CFOP_REVENDA_VALIDOS:
            alertas.append("CFOP {} nao identificado como revenda padrao".format(cfop))
            score += 10

    # 2d. Inconsistencia entre CST PIS e CST COFINS
    if cst_p and cst_c and cst_p != cst_c:
        alertas.append("CST PIS ({}) diverge do CST COFINS ({})".format(cst_pis, cst_cofins))
        score += 20
        log("AVISO", "ValidacaoTributaria",
            "CST PIS diverge de COFINS", "NCM={}".format(ncm))

    score = min(score, 100)

    if not alertas:
        status = "MONOFASICO_VALIDADO"
    elif score >= 50:
        status = "MONOFASICO_COM_RISCO"
        log("RISCO","ValidacaoTributaria",
            "Item com risco tributario alto","NCM={} Score={}".format(ncm,score))
    else:
        status = "MONOFASICO_COM_RISCO"

    return {
        "status_tributario": status,
        "score_risco":       score,
        "motivo_alerta":     " | ".join(alertas) if alertas else "Validado",
    }


# ─────────────────────────────────────────────────────────────
#  SERVICES – leitura XML
# ─────────────────────────────────────────────────────────────
def local_tag(node) -> str:
    t = node.tag
    return t.split("}")[-1] if "}" in t else t

def find_local(el, tag: str):
    for node in el.iter():
        if local_tag(node) == tag:
            return node
    return None

def text_local(el, tag: str, default: str = "") -> str:
    node = find_local(el, tag)
    return node.text.strip() if node is not None and node.text else default

def extrair_data(root) -> str:
    dh = text_local(root, "dhEmi") or text_local(root, "dEmi")
    if not dh: return "SEM-DATA"
    try: return dh[:7]
    except: return "SEM-DATA"

def ler_xml_nfe(nome_arquivo: str, conteudo: bytes) -> tuple[list[dict], str]:
    """
    Le NF-e XML. Retorna (lista_itens, mes_ano).
    Cada item inclui: descricao, ncm, valor, cst_pis, cst_cofins, cfop.
    Robusto a BOM, namespace e estruturas nfeProc/enviNFe.
    """
    if isinstance(conteudo, bytes):
        conteudo = conteudo.lstrip(b"\xef\xbb\xbf").strip()

    # Validacao de tamanho
    tamanho_mb = len(conteudo) / (1024 * 1024)
    if tamanho_mb > UPLOAD_MAX_MB:
        raise ValueError("Arquivo excede limite de {}MB".format(UPLOAD_MAX_MB))

    try:
        root = ET.fromstring(conteudo)
    except ET.ParseError as e:
        raise ValueError("XML malformado: {}".format(e))

    mes_ano = extrair_data(root)
    dets = [n for n in root.iter() if local_tag(n) == "det"]

    if not dets:
        log("AVISO","LeituraXML","Nenhum item encontrado",nome_arquivo)

    itens = []
    for det in dets:
        prod = find_local(det, "prod")
        if prod is None: continue

        descricao = text_local(prod, "xProd")
        ncm_raw   = text_local(prod, "NCM")
        vprod_str = text_local(prod, "vProd")
        cfop      = text_local(prod, "CFOP")
        ncm       = ncm_raw.replace(".", "").replace("-", "").strip()

        # CST PIS/COFINS – extraídos dos grupos PISNT/PISAliq/COFINSNT etc.
        cst_pis    = text_local(det, "CST") if find_local(det, "PIS")    else ""
        cst_cofins = ""
        # Busca especifica para PIS
        pis_node = find_local(det, "PIS")
        if pis_node:
            cst_pis = text_local(pis_node, "CST")
        cof_node = find_local(det, "COFINS")
        if cof_node:
            cst_cofins = text_local(cof_node, "CST")

        try:
            valor = round(float(vprod_str.replace(",",".")), 2)
        except ValueError:
            valor = 0.0
            log("AVISO","LeituraXML","vProd invalido","Arquivo={} NCM={}".format(nome_arquivo,ncm_raw))

        if not ncm_raw:
            log("AVISO","LeituraXML","NCM ausente no item","Arquivo={} Desc={}".format(nome_arquivo,descricao))

        itens.append({
            "descricao":   descricao or "(sem descricao)",
            "ncm_raw":     ncm_raw,
            "ncm":         ncm,
            "valor":       valor,
            "cst_pis":     cst_pis,
            "cst_cofins":  cst_cofins,
            "cfop":        cfop,
        })

    log("INFO","LeituraXML",
        "{} itens extraidos".format(len(itens)), nome_arquivo)
    return itens, mes_ano


# ─────────────────────────────────────────────────────────────
#  SERVICES – leitura PGDAS
# ─────────────────────────────────────────────────────────────
def ler_pgdas(file) -> pd.DataFrame:
    """
    Le CSV ou Excel com colunas: Mes, Receita_PGDAS, DAS_Pago.
    Retorna DataFrame padronizado.
    """
    try:
        nome = file.name.lower()
        if nome.endswith(".csv"):
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
        raise ValueError("Nao foi possivel ler o arquivo PGDAS: {}".format(e))

    df.columns = [c.strip().replace(" ","_") for c in df.columns]
    col_map = {c.lower(): c for c in df.columns}
    esperadas = ["Mes","Receita_PGDAS","DAS_Pago"]
    mapa = {}
    for col in esperadas:
        if col.lower() in col_map:
            mapa[col] = col_map[col.lower()]
        else:
            raise ValueError("Coluna '{}' nao encontrada. Disponíveis: {}".format(col, list(df.columns)))

    df = df.rename(columns={v: k for k, v in mapa.items()})[esperadas].copy()

    def norm_mes(v: str) -> str:
        v = str(v).strip()
        for fmt in ("%Y-%m","%m/%Y","%m-%Y","%Y/%m"):
            try: return datetime.strptime(v, fmt).strftime("%Y-%m")
            except: continue
        return v

    def parse_brl(v: str) -> float:
        s = str(v).strip().replace("R$","").replace(" ","")
        if "," in s and "." in s:
            s = s.replace(".","").replace(",",".")
        elif "," in s:
            s = s.replace(",",".")
        try: return round(float(s), 2)
        except: return 0.0

    df["Mes"]           = df["Mes"].apply(norm_mes)
    df["Receita_PGDAS"] = df["Receita_PGDAS"].apply(parse_brl)
    df["DAS_Pago"]      = df["DAS_Pago"].apply(parse_brl)
    df = df.rename(columns={"Mes":"mes","Receita_PGDAS":"receita_pgdas","DAS_Pago":"das_pago_real"})
    return df.sort_values("mes").reset_index(drop=True)


# ─────────────────────────────────────────────────────────────
#  SERVICES – classificacao NCM
# ─────────────────────────────────────────────────────────────
def classificar_ncm(ncm: str) -> tuple[str, str]:
    ncm_l = str(ncm).strip()
    if not ncm_l:
        return "INCONSISTENCIA", "NCM ausente"
    if not ncm_l.isdigit():
        return "INCONSISTENCIA", "NCM invalido: {}".format(ncm_l)
    for tam in (8, 6, 4):
        chave = ncm_l[:tam].ljust(8, "0")
        if chave in TABELA_NCM_MONOFASICO:
            return "MONOFASICO", TABELA_NCM_MONOFASICO[chave]
    return "NAO MONOFASICO", "NCM fora da tabela"


# ─────────────────────────────────────────────────────────────
#  SERVICES – agrupamento mensal
# ─────────────────────────────────────────────────────────────
def agrupar_por_mes(itens: list[dict]) -> list[dict]:
    meses: dict[str, dict] = {}
    for item in itens:
        m = item.get("mes","SEM-DATA")
        if m not in meses:
            meses[m] = {"total":0.0,"mono":0.0,"mono_validado":0.0,"mono_risco":0.0}
        meses[m]["total"] += item["valor"]
        st = item.get("status_tributario","")
        if st == "MONOFASICO_VALIDADO":
            meses[m]["mono"]          += item["valor"]
            meses[m]["mono_validado"] += item["valor"]
        elif st == "MONOFASICO_COM_RISCO":
            meses[m]["mono"]       += item["valor"]
            meses[m]["mono_risco"] += item["valor"]

    resultado = []
    for m in sorted((k for k in meses if k != "SEM-DATA")) + [k for k in meses if k == "SEM-DATA"]:
        d = meses[m]
        total = round(d["total"], 2)
        mono  = round(d["mono"],  2)
        resultado.append({
            "mes":                m,
            "receita_total":      total,
            "receita_monofasica": mono,
            "receita_mono_valid": round(d["mono_validado"], 2),
            "receita_mono_risco": round(d["mono_risco"], 2),
            "receita_tributavel": round(total - mono, 2),
        })
    return resultado


# ─────────────────────────────────────────────────────────────
#  SERVICES – RBT12 dinamico (fonte PGDAS ou XML)
# ─────────────────────────────────────────────────────────────
def calcular_rbt12(
    agrupamento: list[dict],
    rbt12_inicial: float = 0.0,
    df_pgdas: Optional[pd.DataFrame] = None,
) -> dict[str, float]:
    """
    RBT12 rolling de 12 meses.
    Fonte primaria: Receita_PGDAS (se disponivel).
    Fallback: receita_total dos XMLs.
    """
    pgdas_rec: dict[str, float] = {}
    if df_pgdas is not None and not df_pgdas.empty:
        pgdas_rec = dict(zip(df_pgdas["mes"], df_pgdas["receita_pgdas"]))

    resultado: dict[str, float] = {}
    historico: list[tuple[str, float]] = []
    saldo = rbt12_inicial

    for row in agrupamento:
        mes = row["mes"]
        resultado[mes] = round(saldo, 2)
        rec = pgdas_rec.get(mes, row["receita_total"])
        historico.append((mes, rec))
        saldo += rec
        if len(historico) > 12:
            _, antiga = historico.pop(0)
            saldo -= antiga

    return resultado


# ─────────────────────────────────────────────────────────────
#  SERVICES – faixa e aliquota efetiva
# ─────────────────────────────────────────────────────────────
def identificar_faixa(rbt12: float, tabela: list[dict]) -> Optional[dict]:
    if rbt12 <= 0:
        return tabela[0]
    for f in tabela:
        if rbt12 <= f["limite"]:
            return f
    return None

def calcular_aliquota_efetiva(rbt12: float, faixa: dict) -> float:
    if rbt12 <= 0: return 0.0
    return (rbt12 * faixa["aliquota"] - faixa["deducao"]) / rbt12

def pct_pis_cofins(faixa_n: int, nome_anexo: str) -> float:
    rep = REPARTICAO.get(nome_anexo, {})
    f   = rep.get(faixa_n, {})
    return f.get("pis", 0.0) + f.get("cofins", 0.0)


# ─────────────────────────────────────────────────────────────
#  SERVICES – cruzamento XML vs PGDAS
# ─────────────────────────────────────────────────────────────
def cruzar_xml_pgdas(agrupamento: list[dict], df_pgdas: Optional[pd.DataFrame]) -> list[dict]:
    pgdas_idx: dict[str, dict] = {}
    if df_pgdas is not None and not df_pgdas.empty:
        for _, row in df_pgdas.iterrows():
            pgdas_idx[str(row["mes"])] = row

    resultado = []
    for row in agrupamento:
        pr = pgdas_idx.get(row["mes"])
        if pr is not None:
            rec_p  = round(float(pr["receita_pgdas"]), 2)
            das_r  = round(float(pr["das_pago_real"]), 2)
            div    = round(abs(row["receita_total"] - rec_p) / rec_p, 4) if rec_p > 0 else 0.0
            alerta = div > DIVERGENCIA_MAX
            if alerta:
                log("AVISO","Cruzamento",
                    "Divergencia de {:.1f}% em {}".format(div*100, row["mes"]),
                    "XML={} PGDAS={}".format(brl(row["receita_total"]), brl(rec_p)))
            resultado.append({**row,
                "receita_pgdas":      rec_p,
                "das_pago_real":      das_r,
                "divergencia_pct":    div,
                "tem_pgdas":          True,
                "alerta_divergencia": alerta})
        else:
            if pgdas_idx:  # PGDAS carregado mas mes ausente
                log("INFO","Cruzamento","Mes sem PGDAS",row["mes"])
            resultado.append({**row,
                "receita_pgdas":      None,
                "das_pago_real":      None,
                "divergencia_pct":    None,
                "tem_pgdas":          False,
                "alerta_divergencia": False})
    return resultado


# ─────────────────────────────────────────────────────────────
#  SERVICES – tratamento de credito negativo
# ─────────────────────────────────────────────────────────────
def tratar_credito(credito_bruto: Optional[float], credito_real: Optional[float]) -> dict:
    """
    Classifica o credito como RECUPERAVEL, NAO_RECUPERAVEL ou RISCO_FISCAL.
    Credito negativo = RISCO_FISCAL (nao recuperavel, gera alerta).
    """
    if credito_bruto is None or credito_real is None:
        return {"status_credito":"NAO_RECUPERAVEL","credito_final":None,
                "motivo_credito":"Calculo indisponivel"}
    if credito_bruto < 0:
        log("RISCO","Credito","Credito bruto negativo",
            "Valor={}".format(brl(credito_bruto)))
        return {"status_credito":"RISCO_FISCAL","credito_final":0.0,
                "motivo_credito":"DAS Correto > DAS Usado (verificar parametros)"}
    if credito_real <= 0:
        return {"status_credito":"NAO_RECUPERAVEL","credito_final":0.0,
                "motivo_credito":"Credito real zero ou negativo"}
    return {"status_credito":"RECUPERAVEL","credito_final":round(credito_real,2),
            "motivo_credito":""}


# ─────────────────────────────────────────────────────────────
#  SERVICES – apuracao principal
# ─────────────────────────────────────────────────────────────
def apurar_periodo_real(
    agrupamento: list[dict],
    nome_anexo: str,
    rbt12_inicial: float = 0.0,
    df_pgdas: Optional[pd.DataFrame] = None,
) -> list[dict]:
    """
    Pipeline completo de apuracao mensal com DAS real/estimado,
    RBT12 dinamico, validacao de credito e arredondamento contabil.
    """
    tabela       = TABELAS_SIMPLES[nome_anexo]
    rbt12_map    = calcular_rbt12(agrupamento, rbt12_inicial, df_pgdas)
    cruzado      = cruzar_xml_pgdas(agrupamento, df_pgdas)
    resultado    = []

    for row in cruzado:
        mes   = row["mes"]
        rbt12 = rbt12_map.get(mes, 0.0)
        faixa = identificar_faixa(rbt12, tabela)

        if faixa is None:
            resultado.append({**row,"rbt12":rbt12,"faixa":"ACIMA",
                "aliquota_nominal":None,"aliquota_efetiva":None,
                "das_estimado":None,"das_usado":None,"fonte_das":"–",
                "das_correto":None,"credito_bruto":None,
                "pct_pis_cofins_val":None,"credito_real":None,
                "status_credito":"NAO_RECUPERAVEL",
                "credito_final":0.0,
                "motivo_credito":"RBT12 acima de R$ 4,8M",
                "alerta":"RBT12 acima de R$ 4,8M – fora do Simples"})
            log("RISCO","Apuracao","RBT12 fora do Simples","Mes={}".format(mes))
            continue

        aliq_ef      = calcular_aliquota_efetiva(rbt12, faixa)
        das_est      = r2(row["receita_total"]      * aliq_ef)
        das_correto  = r2(row["receita_tributavel"] * aliq_ef)

        if row["tem_pgdas"] and row["das_pago_real"] is not None:
            das_usado = r2(row["das_pago_real"])
            fonte     = "REAL"
        else:
            das_usado = das_est
            fonte     = "ESTIMADO"

        cred_bruto = r2(das_usado - das_correto) if das_usado is not None else None
        pct_pc     = pct_pis_cofins(faixa["faixa"], nome_anexo)
        cred_real  = r2(cred_bruto * pct_pc) if cred_bruto is not None else None

        status_cred = tratar_credito(cred_bruto, cred_real)

        resultado.append({
            **row,
            "rbt12":              rbt12,
            "faixa":              faixa["faixa"],
            "aliquota_nominal":   faixa["aliquota"],
            "aliquota_efetiva":   aliq_ef,
            "das_estimado":       das_est,
            "das_usado":          das_usado,
            "fonte_das":          fonte,
            "das_correto":        das_correto,
            "credito_bruto":      cred_bruto,
            "pct_pis_cofins_val": pct_pc,
            "credito_real":       cred_real,
            **status_cred,
            "alerta":             "",
        })

    return resultado


# ─────────────────────────────────────────────────────────────
#  SERVICES – resumo geral + score
# ─────────────────────────────────────────────────────────────
def calcular_resumo(itens: list[dict], aliquota: float = ALIQ_ESTIMATIVA) -> dict:
    total       = sum(i["valor"] for i in itens)
    mono        = sum(i["valor"] for i in itens if "MONOFASICO" in i.get("status_tributario",""))
    nao_mono    = sum(i["valor"] for i in itens if i.get("status_tributario") == "NAO_MONOFASICO")
    inconsist   = sum(i["valor"] for i in itens if i.get("status_tributario") == "INCONSISTENTE")
    validado    = sum(i["valor"] for i in itens if i.get("status_tributario") == "MONOFASICO_VALIDADO")
    risco       = sum(i["valor"] for i in itens if i.get("status_tributario") == "MONOFASICO_COM_RISCO")
    return {
        "total_geral":      round(total, 2),
        "total_monofasico": round(mono,  2),
        "total_nao_mono":   round(nao_mono, 2),
        "total_inconsist":  round(inconsist, 2),
        "total_validado":   round(validado, 2),
        "total_risco":      round(risco, 2),
        "estimativa_recup": round(mono * aliquota, 2),
        "pct_monofasico":   round(mono / total, 4) if total > 0 else 0.0,
    }


def calcular_score_oportunidade(
    resumo: dict,
    apuracao: list[dict],
) -> dict:
    """
    Score BAIXA / MEDIA / ALTA baseado em 4 fatores:
      1. Percentual monofasico (40 pts)
      2. Credito potencial absoluto (30 pts)
      3. Consistencia tributaria (20 pts)
      4. Ausencia de divergencias (10 pts)
    """
    score = 0

    # 1. Percentual monofasico
    pct_m = resumo.get("pct_monofasico", 0.0)
    if   pct_m >= 0.60: score += 40
    elif pct_m >= 0.30: score += 25
    elif pct_m >= 0.10: score += 10

    # 2. Credito potencial
    cred = sum(r.get("credito_final", 0) or 0 for r in apuracao)
    if   cred >= 10_000: score += 30
    elif cred >= 3_000:  score += 20
    elif cred >= 500:    score += 10

    # 3. Consistencia (menos itens com risco = melhor)
    n_risco = sum(1 for r in apuracao if r.get("status_credito") == "RISCO_FISCAL")
    n_total = len(apuracao)
    if n_total > 0:
        taxa_risco = n_risco / n_total
        if   taxa_risco == 0.0: score += 20
        elif taxa_risco < 0.2:  score += 10

    # 4. Divergencias
    n_div = sum(1 for r in apuracao if r.get("alerta_divergencia"))
    if   n_div == 0: score += 10
    elif n_div <= 2: score += 5

    if   score >= 70: nivel = "ALTA";  emoji = "🟢"; cor = "green"
    elif score >= 40: nivel = "MEDIA"; emoji = "🟡"; cor = "orange"
    else:             nivel = "BAIXA"; emoji = "🔴"; cor = "red"

    return {"score": score, "nivel": nivel, "emoji": emoji, "cor": cor,
            "credito_total": round(cred, 2)}


def gerar_resumo_executivo(resumo: dict, apuracao: list[dict], score: dict, n_xmls: int) -> str:
    """Gera texto de resumo executivo automatico."""
    cred    = score["credito_total"]
    pct_m   = resumo["pct_monofasico"] * 100
    n_meses = len([r for r in apuracao if r.get("credito_final", 0) and r["credito_final"] > 0])
    n_div   = sum(1 for r in apuracao if r.get("alerta_divergencia"))
    n_risco = sum(1 for r in apuracao if r.get("status_credito") == "RISCO_FISCAL")

    linhas = [
        "Foram analisadas {} nota(s) fiscal(is) abrangendo {} mes(es) de faturamento.".format(
            n_xmls, len(apuracao)),
        "",
        "OPORTUNIDADE IDENTIFICADA",
        "Receita monofasica representa {:.1f}% do faturamento total analisado ({}).".format(
            pct_m, brl(resumo["total_monofasico"])),
        "Credito potencial de PIS/COFINS estimado em {} pela segregacao de receitas monofasicas.".format(
            brl(cred)),
        "Meses com credito recuperavel: {}.".format(n_meses),
        "",
        "VALIDACAO TRIBUTARIA",
        "Receita monofasica validada (sem alertas): {}.".format(brl(resumo["total_validado"])),
        "Receita monofasica com risco tributario: {}.".format(brl(resumo["total_risco"])),
        "",
        "ALERTAS",
        "Divergencias receita XML vs PGDAS (>5%): {} mes(es).".format(n_div),
        "Meses com risco fiscal (credito negativo): {} mes(es).".format(n_risco),
        "",
        "SCORE DE OPORTUNIDADE: {} {} ({}/100)".format(
            score["nivel"], score["emoji"], score["score"]),
        "",
        "Nota: Valores estimados. Recomenda-se validacao por contador ou consultor tributario habilitado.",
        "Base legal: LC 123/2006 · Resolucao CGSN 140/2018.",
    ]
    return "\n".join(linhas)


# ─────────────────────────────────────────────────────────────
#  EXPORTS – Excel profissional com 5 abas
# ─────────────────────────────────────────────────────────────
def gerar_excel_profissional(
    itens: list[dict],
    resumo: dict,
    apuracao: list[dict],
    score: dict,
    texto_resumo: str,
    logs: list[dict],
) -> bytes:
    """
    Gera Excel com 5 abas:
    1. Resumo Executivo
    2. Apuracao Mensal
    3. Itens Classificados
    4. Inconsistencias Tributarias
    5. Logs e Alertas
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as w:

        # ABA 1 – Resumo Executivo
        df_exec = pd.DataFrame([
            {"Indicador": "Faturamento Total",           "Valor": brl(resumo["total_geral"])},
            {"Indicador": "Faturamento Monofasico",      "Valor": brl(resumo["total_monofasico"])},
            {"Indicador": "  Validado (sem alertas)",    "Valor": brl(resumo["total_validado"])},
            {"Indicador": "  Com risco tributario",      "Valor": brl(resumo["total_risco"])},
            {"Indicador": "Faturamento Nao Monofasico",  "Valor": brl(resumo["total_nao_mono"])},
            {"Indicador": "% Monofasico / Total",        "Valor": pct_fmt(resumo["pct_monofasico"])},
            {"Indicador": "Credito Potencial Total",     "Valor": brl(score["credito_total"])},
            {"Indicador": "Score de Oportunidade",       "Valor": "{} ({}/100)".format(score["nivel"],score["score"])},
            {"Indicador": "Divergencias encontradas",    "Valor": str(sum(1 for r in apuracao if r.get("alerta_divergencia")))},
            {"Indicador": "Meses com risco fiscal",      "Valor": str(sum(1 for r in apuracao if r.get("status_credito")=="RISCO_FISCAL"))},
        ])
        df_exec.to_excel(w, sheet_name="1.Resumo Executivo", index=False)
        pd.DataFrame([{"Texto": l} for l in texto_resumo.split("\n")]).to_excel(
            w, sheet_name="1.Resumo Executivo", index=False, startrow=len(df_exec)+2)

        # ABA 2 – Apuracao Mensal
        linhas_ap = []
        for r in apuracao:
            linhas_ap.append({
                "Mes":            fmt_mes(r["mes"]),
                "RBT12":          brl(r.get("rbt12")),
                "Faixa":          r.get("faixa","–"),
                "Aliq. Nominal":  pct_fmt(r.get("aliquota_nominal")),
                "Aliq. Efetiva":  pct_fmt(r.get("aliquota_efetiva"),4),
                "Rec. XML":       brl(r.get("receita_total")),
                "Rec. PGDAS":     brl(r.get("receita_pgdas")),
                "Divergencia":    pct_fmt(r.get("divergencia_pct")) if r.get("divergencia_pct") is not None else "–",
                "Rec. Monof.":    brl(r.get("receita_monofasica")),
                "Rec. Validada":  brl(r.get("receita_mono_valid")),
                "Rec. c/ Risco":  brl(r.get("receita_mono_risco")),
                "Rec. Tributavel":brl(r.get("receita_tributavel")),
                "DAS Usado":      brl(r.get("das_usado")),
                "Fonte DAS":      r.get("fonte_das","–"),
                "DAS Correto":    brl(r.get("das_correto")),
                "Cred. Bruto":    brl(r.get("credito_bruto")),
                "% PIS+COF":      pct_fmt(r.get("pct_pis_cofins_val")),
                "Cred. Real":     brl(r.get("credito_real")),
                "Cred. Final":    brl(r.get("credito_final")),
                "Status Credito": r.get("status_credito","–"),
                "Motivo":         r.get("motivo_credito",""),
            })
        pd.DataFrame(linhas_ap).to_excel(w, sheet_name="2.Apuracao Mensal", index=False)

        # ABA 3 – Itens Classificados
        linhas_it = [{
            "Mes":            i.get("mes",""),
            "Arquivo":        i.get("arquivo",""),
            "Descricao":      i.get("descricao",""),
            "NCM":            i.get("ncm",""),
            "CFOP":           i.get("cfop",""),
            "CST PIS":        i.get("cst_pis",""),
            "CST COFINS":     i.get("cst_cofins",""),
            "Valor (R$)":     brl(i.get("valor")),
            "Class. NCM":     i.get("classificacao",""),
            "Status Trib.":   i.get("status_tributario",""),
            "Score Risco":    i.get("score_risco",""),
            "Motivo Alerta":  i.get("motivo_alerta",""),
        } for i in itens]
        pd.DataFrame(linhas_it).to_excel(w, sheet_name="3.Itens Classificados", index=False)

        # ABA 4 – Inconsistencias
        incons = [i for i in itens if i.get("status_tributario") in
                  ("INCONSISTENTE","MONOFASICO_COM_RISCO")]
        if incons:
            pd.DataFrame([{
                "Mes":           i.get("mes",""),
                "Descricao":     i.get("descricao",""),
                "NCM":           i.get("ncm",""),
                "CFOP":          i.get("cfop",""),
                "CST PIS":       i.get("cst_pis",""),
                "CST COFINS":    i.get("cst_cofins",""),
                "Valor":         brl(i.get("valor")),
                "Status":        i.get("status_tributario",""),
                "Score":         i.get("score_risco",""),
                "Motivo":        i.get("motivo_alerta",""),
            } for i in incons]).to_excel(w, sheet_name="4.Inconsistencias", index=False)
        else:
            pd.DataFrame([{"Info":"Nenhuma inconsistencia encontrada"}]).to_excel(
                w, sheet_name="4.Inconsistencias", index=False)

        # ABA 5 – Logs
        pd.DataFrame(logs).to_excel(w, sheet_name="5.Logs e Alertas", index=False)

    return output.getvalue()


def gerar_csv_apuracao(apuracao: list[dict]) -> bytes:
    linhas = [{
        "Mes":           r.get("mes",""),
        "RBT12":         r.get("rbt12",""),
        "Faixa":         r.get("faixa",""),
        "Aliq Efetiva":  "{:.4f}%".format(r["aliquota_efetiva"]*100) if r.get("aliquota_efetiva") is not None else "",
        "Rec XML":       r.get("receita_total",""),
        "Rec PGDAS":     r.get("receita_pgdas","") or "",
        "Divergencia":   "{:.2f}%".format(r["divergencia_pct"]*100) if r.get("divergencia_pct") is not None else "",
        "Rec Monof":     r.get("receita_monofasica",""),
        "Rec Tributavel":r.get("receita_tributavel",""),
        "DAS Usado":     r.get("das_usado",""),
        "Fonte DAS":     r.get("fonte_das",""),
        "DAS Correto":   r.get("das_correto",""),
        "Cred Bruto":    r.get("credito_bruto",""),
        "Pct PIS+COF":   "{:.2f}%".format(r["pct_pis_cofins_val"]*100) if r.get("pct_pis_cofins_val") is not None else "",
        "Cred Real":     r.get("credito_real",""),
        "Cred Final":    r.get("credito_final",""),
        "Status Credito":r.get("status_credito",""),
    } for r in apuracao]
    return pd.DataFrame(linhas).to_csv(index=False, sep=";", decimal=",").encode("utf-8")


# ─────────────────────────────────────────────────────────────
#  PIPELINE PRINCIPAL
# ─────────────────────────────────────────────────────────────
def processar_xmls(arquivos: list[tuple[str,bytes]], aliquota: float) -> tuple[list[dict], dict]:
    todos: list[dict] = []
    for nome, conteudo in arquivos:
        try:
            itens_raw, mes_ano = ler_xml_nfe(nome, conteudo)
        except ValueError as e:
            st.warning("Erro ao ler '{}': {}".format(nome, e))
            log("ERRO","LeituraXML",str(e),nome)
            continue
        for item in itens_raw:
            classif, motivo = classificar_ncm(item["ncm"])
            val_trib = validar_item_tributario(
                item["ncm"], item["cst_pis"], item["cst_cofins"],
                item["cfop"], classif)
            todos.append({
                "arquivo":          nome,
                "mes":              mes_ano,
                "descricao":        item["descricao"],
                "ncm":              item["ncm_raw"],
                "cfop":             item["cfop"],
                "cst_pis":          item["cst_pis"],
                "cst_cofins":       item["cst_cofins"],
                "valor":            item["valor"],
                "classificacao":    classif,
                "motivo_ncm":       motivo,
                **val_trib,
            })
    return todos, calcular_resumo(todos, aliquota)


# ─────────────────────────────────────────────────────────────
#  INTERFACE STREAMLIT
# ─────────────────────────────────────────────────────────────
def main() -> None:
    st.set_page_config(
        page_title="PIS/COFINS Pro",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    # CSS customizado – visual SaaS
    st.markdown("""
    <style>
    .metric-card {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d5a8e 100%);
        border-radius:12px; padding:20px; margin:4px;
        color:white; text-align:center;
    }
    .metric-card h3 { font-size:14px; opacity:0.8; margin:0; }
    .metric-card h1 { font-size:24px; font-weight:700; margin:4px 0; }
    .badge-green  { background:#1a7a4a; color:white; padding:3px 10px; border-radius:12px; font-size:12px; }
    .badge-yellow { background:#b8860b; color:white; padding:3px 10px; border-radius:12px; font-size:12px; }
    .badge-red    { background:#8b1a1a; color:white; padding:3px 10px; border-radius:12px; font-size:12px; }
    .badge-blue   { background:#1a4a8b; color:white; padding:3px 10px; border-radius:12px; font-size:12px; }
    .section-header { font-size:20px; font-weight:700; color:#1e3a5f;
                      border-left:4px solid #2d5a8e; padding-left:12px; margin:20px 0 10px 0; }
    div[data-testid="stMetricValue"] { font-size:22px !important; }
    </style>
    """, unsafe_allow_html=True)

    # Header
    col_t, col_v = st.columns([6,1])
    with col_t:
        st.markdown("## 📊 PIS/COFINS Pro – Recuperacao Tributaria")
        st.caption("Simples Nacional · Regime de Revenda · Analise por NCM/CST/CFOP · v{}".format(VERSAO))
    with col_v:
        st.markdown("<br>", unsafe_allow_html=True)
        st.caption("v{}".format(VERSAO))
    st.divider()

    limpar_logs()

    # ── SIDEBAR ───────────────────────────────────────────────
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/brazil.png", width=48)
        st.markdown("### Configuracoes")

        nome_anexo = st.selectbox("Anexo do Simples", list(TABELAS_SIMPLES.keys()))
        rbt12_ini  = st.number_input("RBT12 anterior ao periodo (R$)",
                                     min_value=0.0, value=360_000.0,
                                     step=1_000.0, format="%.2f",
                                     help="Receita dos 12 meses anteriores ao primeiro XML.")
        usar_pgdas = st.checkbox("Usar dados reais do PGDAS", value=False)
        aliq_pct   = st.number_input("Aliquota estimativa PIS+COFINS (%)",
                                     min_value=0.0, max_value=100.0,
                                     value=9.25, step=0.05, format="%.2f")
        aliq_dec   = aliq_pct / 100

        st.markdown("---")
        with st.expander("Reparticao do Anexo"):
            rep = REPARTICAO.get(nome_anexo, {})
            st.dataframe(pd.DataFrame([{
                "Faixa":      fn,
                "PIS":        pct_fmt(v.get("pis",0)),
                "COFINS":     pct_fmt(v.get("cofins",0)),
                "PIS+COFINS": pct_fmt(v.get("pis",0)+v.get("cofins",0)),
            } for fn,v in rep.items()]), hide_index=True, use_container_width=True)

        st.markdown("---")
        st.caption("`{}` NCMs na tabela".format(len(TABELA_NCM_MONOFASICO)))

    # ── UPLOAD XMLs ───────────────────────────────────────────
    st.markdown('<div class="section-header">Upload de Documentos Fiscais</div>', unsafe_allow_html=True)

    col_xml, col_pgdas = st.columns([1,1]) if usar_pgdas else (st.columns([1,1]))

    with col_xml:
        st.markdown("**XMLs de NF-e**")
        uploaded_xmls = st.file_uploader(
            "Arraste ou selecione os XMLs",
            type=["xml"], accept_multiple_files=True, key="xmls",
            help="Limite: {}MB por arquivo".format(UPLOAD_MAX_MB))

    df_pgdas_df = None
    with col_pgdas:
        if usar_pgdas:
            st.markdown("**Planilha PGDAS/DAS**")
            with st.expander("Ver formato esperado"):
                st.markdown("""
| Mes | Receita_PGDAS | DAS_Pago |
|---|---|---|
| 2024-01 | 45000,00 | 1980,00 |
                """)
                st.download_button("Baixar template",
                    data="Mes;Receita_PGDAS;DAS_Pago\n2024-01;45000,00;1980,00\n".encode(),
                    file_name="template_pgdas.csv", mime="text/csv")
            up_pgdas = st.file_uploader("CSV ou Excel do PGDAS",
                                        type=["csv","xlsx","xls"], key="pgdas")
            if up_pgdas:
                try:
                    df_pgdas_df = ler_pgdas(up_pgdas)
                    st.success("{} mes(es) carregados ({} a {})".format(
                        len(df_pgdas_df),
                        fmt_mes(df_pgdas_df["mes"].min()),
                        fmt_mes(df_pgdas_df["mes"].max())))
                except ValueError as e:
                    st.error(str(e))
        else:
            st.info("Ative 'Usar dados reais do PGDAS' na barra lateral para carregar a planilha do PGDAS.")

    if not uploaded_xmls:
        st.markdown("---")
        st.info("📂 Aguardando upload dos XMLs de NF-e para iniciar a analise.")
        st.stop()

    # ── PROCESSAMENTO ─────────────────────────────────────────
    prog = st.progress(0, text="Iniciando processamento...")
    arquivos = [(f.name, f.read()) for f in uploaded_xmls]
    prog.progress(20, text="Lendo e classificando XMLs...")
    itens, resumo = processar_xmls(arquivos, aliq_dec)
    prog.progress(60, text="Apurando periodo...")

    if not itens:
        st.error("Nenhum item extraido dos XMLs. Verifique os arquivos.")
        st.stop()

    agrupado  = agrupar_por_mes(itens)
    apuracao  = apurar_periodo_real(agrupado, nome_anexo, rbt12_ini, df_pgdas_df)
    score     = calcular_score_oportunidade(resumo, apuracao)
    texto_ex  = gerar_resumo_executivo(resumo, apuracao, score, len(uploaded_xmls))
    prog.progress(90, text="Gerando dashboard...")
    df_itens  = pd.DataFrame(itens)
    prog.progress(100, text="Concluido!")
    prog.empty()

    # ── SCORE DE OPORTUNIDADE ─────────────────────────────────
    st.markdown('<div class="section-header">Score de Oportunidade</div>', unsafe_allow_html=True)
    s_col1, s_col2, s_col3 = st.columns([1,2,1])
    with s_col2:
        cor_map = {"ALTA":"#1a7a4a","MEDIA":"#b8860b","BAIXA":"#8b1a1a"}
        st.markdown("""
        <div style="background:{cor};border-radius:16px;padding:24px;text-align:center;color:white;">
            <h2 style="margin:0;font-size:48px;">{emoji}</h2>
            <h1 style="margin:8px 0;font-size:36px;font-weight:800;">{nivel} OPORTUNIDADE</h1>
            <p style="font-size:20px;margin:0;opacity:0.9;">Score: {score}/100</p>
            <p style="font-size:16px;margin:8px 0 0 0;opacity:0.8;">Credito Potencial: {cred}</p>
        </div>
        """.format(
            cor=cor_map[score["nivel"]],
            emoji=score["emoji"],
            nivel=score["nivel"],
            score=score["score"],
            cred=brl(score["credito_total"]),
        ), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ── METRICAS GERAIS ───────────────────────────────────────
    st.markdown('<div class="section-header">Resumo Geral</div>', unsafe_allow_html=True)
    m1,m2,m3,m4,m5 = st.columns(5)
    m1.metric("Faturamento Total",       brl(resumo["total_geral"]))
    m2.metric("Receita Monofasica",      brl(resumo["total_monofasico"]),
              delta=pct_fmt(resumo["pct_monofasico"])+" do total")
    m3.metric("Monof. Validado",         brl(resumo["total_validado"]))
    m4.metric("Monof. c/ Risco",         brl(resumo["total_risco"]))
    m5.metric("Credito Potencial",       brl(score["credito_total"]))

    # ── DASHBOARD ─────────────────────────────────────────────
    st.markdown('<div class="section-header">Dashboard Executivo</div>', unsafe_allow_html=True)
    tab_graf1, tab_graf2, tab_graf3 = st.tabs(["Composicao", "Evolucao Mensal", "Ranking NCM"])

    with tab_graf1:
        gc1, gc2 = st.columns(2)
        with gc1:
            fig_pie = px.pie(
                names=["Validado","Com Risco","Nao Monofasico","Inconsistente"],
                values=[resumo["total_validado"], resumo["total_risco"],
                        resumo["total_nao_mono"],  resumo["total_inconsist"]],
                title="Composicao do Faturamento",
                color_discrete_sequence=["#1a7a4a","#b8860b","#2d5a8e","#8b1a1a"],
            )
            st.plotly_chart(fig_pie, use_container_width=True)
        with gc2:
            valid = [r for r in apuracao if r.get("credito_final") is not None]
            if valid:
                fig_cred = px.bar(
                    x=[fmt_mes(r["mes"]) for r in valid],
                    y=[r.get("credito_final",0) or 0 for r in valid],
                    title="Credito Final por Mes (R$)",
                    labels={"x":"Mes","y":"R$"},
                    color=[r.get("status_credito","") for r in valid],
                    color_discrete_map={"RECUPERAVEL":"#1a7a4a",
                                        "NAO_RECUPERAVEL":"#2d5a8e",
                                        "RISCO_FISCAL":"#8b1a1a"},
                )
                st.plotly_chart(fig_cred, use_container_width=True)

    with tab_graf2:
        if agrupado:
            df_ev = pd.DataFrame([{
                "Mes":        fmt_mes(r["mes"]),
                "Total":      r["receita_total"],
                "Monofasico": r["receita_monofasica"],
                "Tributavel": r["receita_tributavel"],
            } for r in agrupado])
            fig_line = px.line(df_ev, x="Mes",
                               y=["Total","Monofasico","Tributavel"],
                               title="Evolucao Mensal da Receita",
                               markers=True)
            st.plotly_chart(fig_line, use_container_width=True)

    with tab_graf3:
        df_ncm = df_itens[df_itens["classificacao"]=="MONOFASICO"].groupby("ncm")["valor"].sum().reset_index()
        df_ncm = df_ncm.sort_values("valor", ascending=False).head(15)
        if not df_ncm.empty:
            fig_ncm = px.bar(df_ncm, x="valor", y="ncm", orientation="h",
                             title="Top 15 NCMs Monofasicos por Valor (R$)",
                             labels={"valor":"R$","ncm":"NCM"})
            st.plotly_chart(fig_ncm, use_container_width=True)

    # ── APURACAO MENSAL ───────────────────────────────────────
    st.markdown('<div class="section-header">Apuracao PGDAS-D – Credito Real PIS/COFINS</div>', unsafe_allow_html=True)

    # Alertas
    for r in apuracao:
        if r.get("alerta"):
            st.error("🚫 {}: {}".format(fmt_mes(r["mes"]), r["alerta"]))
        if r.get("alerta_divergencia"):
            st.warning("⚠️ Divergencia em {}: XML {} vs PGDAS {} ({})".format(
                fmt_mes(r["mes"]), brl(r["receita_total"]),
                brl(r.get("receita_pgdas")),
                pct_fmt(r.get("divergencia_pct"))))
        if r.get("status_credito") == "RISCO_FISCAL":
            st.error("🔴 Risco Fiscal em {}: {}".format(
                fmt_mes(r["mes"]), r.get("motivo_credito","")))

    linhas_ap = []
    for r in apuracao:
        l = {
            "Mes":           fmt_mes(r["mes"]),
            "RBT12":         brl(r.get("rbt12")),
            "Faixa":         str(r.get("faixa","–")),
            "Aliq. Ef.":     pct_fmt(r.get("aliquota_efetiva"),4),
            "Rec. XML":      brl(r.get("receita_total")),
        }
        if usar_pgdas:
            l["Rec. PGDAS"]  = brl(r.get("receita_pgdas"))
            l["Diverg."]     = ("🔴 " if r.get("alerta_divergencia") else "✅ ") + pct_fmt(r.get("divergencia_pct")) if r.get("divergencia_pct") is not None else "–"
        l.update({
            "Rec. Monof.":   brl(r.get("receita_monofasica")),
            "Rec. Trib.":    brl(r.get("receita_tributavel")),
            "DAS Usado":     brl(r.get("das_usado")),
            "Fonte":         ("🟢 REAL" if r.get("fonte_das")=="REAL" else "🟡 EST.") if usar_pgdas else "🟡 EST.",
            "DAS Correto":   brl(r.get("das_correto")),
            "Cred. Bruto":   brl(r.get("credito_bruto")),
            "% PIS+COF":     pct_fmt(r.get("pct_pis_cofins_val")),
            "Cred. Final":   brl(r.get("credito_final")),
            "Status":        {"RECUPERAVEL":"✅","NAO_RECUPERAVEL":"⚪","RISCO_FISCAL":"🔴"}.get(r.get("status_credito",""),"–"),
        })
        linhas_ap.append(l)

    st.dataframe(pd.DataFrame(linhas_ap), use_container_width=True, hide_index=True)

    # Totais
    valid = [r for r in apuracao if r.get("credito_final") is not None]
    t1,t2,t3,t4 = st.columns(4)
    t1.metric("DAS Total Usado",      brl(sum(r.get("das_usado",0) or 0 for r in valid)))
    t2.metric("DAS Total Correto",    brl(sum(r.get("das_correto",0) or 0 for r in valid)))
    t3.metric("Credito Bruto Total",  brl(sum(r.get("credito_bruto",0) or 0 for r in valid)))
    t4.metric("Credito Real PIS/COF", brl(score["credito_total"]),
              delta="a recuperar" if score["credito_total"] > 0 else None)

    if usar_pgdas:
        n_real = sum(1 for r in valid if r.get("fonte_das")=="REAL")
        n_est  = sum(1 for r in valid if r.get("fonte_das")=="ESTIMADO")
        st.caption("🟢 {} mes(es) com DAS real  |  🟡 {} mes(es) estimado".format(n_real, n_est))

    # ── RESUMO EXECUTIVO ──────────────────────────────────────
    st.markdown('<div class="section-header">Resumo Executivo Automatico</div>', unsafe_allow_html=True)
    with st.expander("Ver Resumo Executivo", expanded=True):
        st.text(texto_ex)

    # ── OPORTUNIDADE E RISCOS ─────────────────────────────────
    col_op, col_ri = st.columns(2)
    with col_op:
        st.markdown('<div class="section-header">Oportunidade Identificada</div>', unsafe_allow_html=True)
        meses_top = sorted(
            [r for r in apuracao if (r.get("credito_final") or 0) > 0],
            key=lambda x: x.get("credito_final",0), reverse=True)[:3]
        if meses_top:
            for r in meses_top:
                st.success("**{}** – Credito: {} | Monof.: {}".format(
                    fmt_mes(r["mes"]), brl(r.get("credito_final")),
                    brl(r.get("receita_monofasica"))))
        else:
            st.info("Nenhum credito recuperavel identificado.")

    with col_ri:
        st.markdown('<div class="section-header">Riscos Identificados</div>', unsafe_allow_html=True)
        riscos_items = [i for i in itens if i.get("status_tributario") == "MONOFASICO_COM_RISCO"]
        if riscos_items:
            for it in riscos_items[:5]:
                st.warning("**{}** | NCM {} | Score {} | {}".format(
                    it["descricao"][:35], it["ncm"],
                    it.get("score_risco",0), it.get("motivo_alerta","")[:60]))
            if len(riscos_items) > 5:
                st.caption("... e mais {} item(ns) com risco.".format(len(riscos_items)-5))
        else:
            st.success("Nenhum risco tributario identificado.")

    # ── ITENS CLASSIFICADOS ───────────────────────────────────
    st.markdown('<div class="section-header">Itens Classificados</div>', unsafe_allow_html=True)
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        filtro_st = st.selectbox("Status Tributario", [
            "Todos","MONOFASICO_VALIDADO","MONOFASICO_COM_RISCO","NAO_MONOFASICO","INCONSISTENTE"])
    with col_f2:
        filtro_mes = st.selectbox("Mes", ["Todos"] + sorted(df_itens["mes"].unique().tolist()))

    df_ex = df_itens.copy()
    if filtro_st != "Todos":
        df_ex = df_ex[df_ex["status_tributario"] == filtro_st]
    if filtro_mes != "Todos":
        df_ex = df_ex[df_ex["mes"] == filtro_mes]

    st.dataframe(
        df_ex[["mes","arquivo","descricao","ncm","cfop","cst_pis","cst_cofins",
               "valor","status_tributario","score_risco","motivo_alerta"]].rename(columns={
            "mes":"Mes","arquivo":"Arquivo","descricao":"Descricao","ncm":"NCM",
            "cfop":"CFOP","cst_pis":"CST PIS","cst_cofins":"CST COFINS",
            "valor":"Valor (R$)","status_tributario":"Status Trib.",
            "score_risco":"Score Risco","motivo_alerta":"Motivo Alerta",
        }),
        use_container_width=True, height=380,
    )

    # ── EXPORTS ───────────────────────────────────────────────
    st.markdown('<div class="section-header">Exportar Relatorios</div>', unsafe_allow_html=True)
    e1, e2 = st.columns(2)
    with e1:
        st.download_button(
            "📥 Relatorio Excel Profissional (5 abas)",
            data=gerar_excel_profissional(itens, resumo, apuracao, score, texto_ex, get_logs()),
            file_name="relatorio_piscofins_pro_{}.xlsx".format(datetime.now().strftime("%Y%m%d_%H%M")),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with e2:
        st.download_button(
            "📥 Apuracao CSV",
            data=gerar_csv_apuracao(apuracao),
            file_name="apuracao_pgdas_{}.csv".format(datetime.now().strftime("%Y%m%d_%H%M")),
            mime="text/csv",
            use_container_width=True,
        )

    # ── LOGS ──────────────────────────────────────────────────
    with st.expander("Logs de Auditoria ({} eventos)".format(len(get_logs()))):
        df_logs = pd.DataFrame(get_logs())
        if not df_logs.empty:
            nivel_cores = {"INFO":"🔵","AVISO":"🟡","ERRO":"🔴","RISCO":"🟠"}
            df_logs["nivel"] = df_logs["nivel"].apply(lambda x: nivel_cores.get(x,"")+x)
            st.dataframe(df_logs, use_container_width=True, hide_index=True)
        else:
            st.success("Nenhum evento registrado.")

    # ── METODOLOGIA ───────────────────────────────────────────
    with st.expander("Metodologia e Base Legal"):
        st.markdown("""
**Classificacao Tributaria:**
NCM consultado na Tabela 4.3.10 (EFD-Contribuicoes SPED).
Validacao cruzada com CST PIS/COFINS e CFOP declarados na NF-e.

**Status Tributario:**
- MONOFASICO_VALIDADO: NCM monofasico, CST e CFOP compativeis
- MONOFASICO_COM_RISCO: NCM monofasico, mas CST/CFOP divergente
- NAO_MONOFASICO: NCM fora da tabela
- INCONSISTENTE: CST indica monofasico mas NCM discorda

**Aliquota Efetiva (PGDAS-D):**
`aliquota_efetiva = (RBT12 x aliquota_nominal - deducao) / RBT12`

**Credito:**
`Cred. Bruto = DAS Usado - DAS Correto`
`Cred. Real  = Cred. Bruto x (% PIS + % COFINS da reparticao)`
Credito negativo = RISCO_FISCAL (nao recuperavel).

**Base Legal:** LC 123/2006 · Resolucao CGSN 140/2018 · Lei 10.147/2000
        """)

    st.divider()
    st.caption("PIS/COFINS Pro v{} · Analise preliminar – validar com contador habilitado · LC 123/2006".format(VERSAO))


if __name__ == "__main__":
    main()

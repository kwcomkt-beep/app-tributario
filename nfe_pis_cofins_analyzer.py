"""
=============================================================
  PIS/COFINS PRO  |  v7.0
  Recuperacao Tributaria · Simples Nacional · Regime Revenda
=============================================================
  Estrutura modular (pronta para separacao em pacotes):
    MODULE: constants      → tabelas, constantes globais
    MODULE: utils          → formatacao, arredondamento, logs
    MODULE: xml_parser     → leitura e extracao de NF-e
    MODULE: pgdas          → leitura e normalizacao do PGDAS
    MODULE: validators     → validacao tributaria NCM/CST/CFOP
    MODULE: apuracao       → calculo DAS, RBT12, credito
    MODULE: auditoria      → motor de auditoria tributaria
    MODULE: exports        → CSV, Excel profissional, PDF
    MODULE: dashboard      → graficos e metricas visuais
    MODULE: ui             → interface Streamlit

  Dependencias: pip install streamlit pandas openpyxl plotly reportlab lxml
=============================================================
"""

from __future__ import annotations

import io
import logging
import os
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from io import BytesIO
from typing import Optional

import numpy as np
import pandas as pd
import streamlit as st

# ── importacao segura de bibliotecas opcionais ────────────────
try:
    import plotly.express as px
    import plotly.graph_objects as go
    PLOTLY_OK = True
except ImportError:
    PLOTLY_OK = False

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
    )
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

try:
    import openpyxl
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

# ─────────────────────────────────────────────────────────────
#  MODULE: constants
# ─────────────────────────────────────────────────────────────
VERSAO            = "7.0"
UPLOAD_MAX_MB     = 50
DIVERGENCIA_MAX   = 0.05
ALIQ_ESTIMATIVA   = 0.0925
SEG_PARCIAL_MIN   = 0.05   # >= 5% = segregacao parcial
SEG_TOTAL_MIN     = 0.90   # >= 90% = segregacao total

CST_MONOFASICO    = {"04", "4"}
CST_NORMAL        = {"01","1","02","2","05","5","06","6","07","7","08","8","09","9"}
CFOP_REVENDA      = {"5102","5403","5405","6102","6403","6404","5101","6101","5104","6104"}
CFOP_INCOMPAT     = {"5201","5202","6201","6202","5910","6910","5949","6949"}

TABELA_NCM: dict[str, str] = {
    "27101112":"Gasolina automotiva comum","27101113":"Gasolina automotiva premium",
    "27101121":"Querosene de aviacao","27101131":"Oleo diesel",
    "27101500":"Oleos lubrificantes","27111100":"GNL","27111910":"GLP",
    "30011000":"Glandulas/orgaos opoterapicos","30021000":"Antissoros/imunoglobulinas",
    "30022000":"Vacinas medicina humana","30023000":"Vacinas medicina veterinaria",
    "30031000":"Medicamentos c/ penicilinas","30032000":"Medicamentos c/ antibioticos",
    "30039099":"Outros medicamentos mistura","30041000":"Medicamentos penicilinas (doses)",
    "30042000":"Medicamentos antibioticos (doses)","30043900":"Medicamentos hormonais",
    "30049099":"Outros medicamentos uso humano","33011000":"Oleos essenciais citricos",
    "33012900":"Outros oleos essenciais","33030010":"Perfumes (extratos)",
    "33030020":"Aguas-de-colonia","33041000":"Maquiagem labios",
    "33042000":"Sombras e delineadores","33049900":"Outros produtos beleza",
    "33051000":"Xampus","33052000":"Preparacoes ondulacao/alisamento",
    "33053000":"Laques","33059000":"Outras preparacoes capilares",
    "33061000":"Dentifricio","33062000":"Fio dental",
    "33069000":"Higiene bucal outros","33071000":"Preparacoes barbear",
    "33072000":"Desodorantes e antiperspirantes","33074900":"Outros toucador",
    "22011000":"Agua mineral/gaseificada","22019000":"Outras aguas",
    "22021000":"Agua c/ acucar/adocante","22029000":"Bebidas nao alcoolicas",
    "22030000":"Cerveja de malte","22060000":"Bebidas fermentadas",
    "22071000":"Alcool etilico >= 80%","22082000":"Conhaque","22083000":"Uisque",
    "22084000":"Rum e tafia","22085000":"Gim e genebra","22086000":"Vodca",
    "22087000":"Licores","22089900":"Outras bebidas alcoolicas",
    "87031000":"Veiculos neve/quadriciclos","87032100":"Automoveis <= 1000 cm3",
    "87032200":"Automoveis 1000-1500 cm3","87032300":"Automoveis 1500-3000 cm3",
    "87032400":"Automoveis > 3000 cm3","87033300":"Automoveis diesel > 2500 cm3",
    "87060010":"Chassis c/ motor","87089900":"Acessorios veiculos",
    "87111000":"Motos <= 50 cm3","87112000":"Motos 50-250 cm3",
    "87113000":"Motos 250-500 cm3","87114000":"Motos 500-800 cm3",
    "87115000":"Motos > 800 cm3","40111000":"Pneus novos automoveis",
    "40112000":"Pneus novos onibus/caminhoes","40113000":"Pneus novos avioes",
    "40114000":"Pneus novos motocicletas","40119100":"Pneus novos outros",
    "40121100":"Pneus recauch. automoveis","40121200":"Pneus recauch. onibus/caminhoes",
}

TABELAS_SIMPLES: dict[str, list[dict]] = {
    "Anexo I – Comercio": [
        {"faixa":1,"limite":180_000,   "aliquota":0.04,  "deducao":0.0},
        {"faixa":2,"limite":360_000,   "aliquota":0.073, "deducao":5_940.0},
        {"faixa":3,"limite":720_000,   "aliquota":0.095, "deducao":13_860.0},
        {"faixa":4,"limite":1_800_000, "aliquota":0.107, "deducao":22_500.0},
        {"faixa":5,"limite":3_600_000, "aliquota":0.143, "deducao":87_300.0},
        {"faixa":6,"limite":4_800_000, "aliquota":0.19,  "deducao":378_000.0},
    ],
    "Anexo II – Industria": [
        {"faixa":1,"limite":180_000,   "aliquota":0.045, "deducao":0.0},
        {"faixa":2,"limite":360_000,   "aliquota":0.078, "deducao":5_940.0},
        {"faixa":3,"limite":720_000,   "aliquota":0.10,  "deducao":13_860.0},
        {"faixa":4,"limite":1_800_000, "aliquota":0.113, "deducao":22_500.0},
        {"faixa":5,"limite":3_600_000, "aliquota":0.147, "deducao":85_500.0},
        {"faixa":6,"limite":4_800_000, "aliquota":0.30,  "deducao":720_000.0},
    ],
    "Anexo III – Servicos A": [
        {"faixa":1,"limite":180_000,   "aliquota":0.06,  "deducao":0.0},
        {"faixa":2,"limite":360_000,   "aliquota":0.112, "deducao":9_360.0},
        {"faixa":3,"limite":720_000,   "aliquota":0.135, "deducao":17_640.0},
        {"faixa":4,"limite":1_800_000, "aliquota":0.16,  "deducao":35_640.0},
        {"faixa":5,"limite":3_600_000, "aliquota":0.21,  "deducao":125_640.0},
        {"faixa":6,"limite":4_800_000, "aliquota":0.33,  "deducao":648_000.0},
    ],
    "Anexo IV – Servicos B": [
        {"faixa":1,"limite":180_000,   "aliquota":0.045, "deducao":0.0},
        {"faixa":2,"limite":360_000,   "aliquota":0.09,  "deducao":8_100.0},
        {"faixa":3,"limite":720_000,   "aliquota":0.102, "deducao":12_420.0},
        {"faixa":4,"limite":1_800_000, "aliquota":0.14,  "deducao":39_780.0},
        {"faixa":5,"limite":3_600_000, "aliquota":0.22,  "deducao":183_780.0},
        {"faixa":6,"limite":4_800_000, "aliquota":0.33,  "deducao":828_000.0},
    ],
    "Anexo V – Servicos C": [
        {"faixa":1,"limite":180_000,   "aliquota":0.15,  "deducao":0.0},
        {"faixa":2,"limite":360_000,   "aliquota":0.18,  "deducao":5_400.0},
        {"faixa":3,"limite":720_000,   "aliquota":0.195, "deducao":13_500.0},
        {"faixa":4,"limite":1_800_000, "aliquota":0.205, "deducao":20_700.0},
        {"faixa":5,"limite":3_600_000, "aliquota":0.23,  "deducao":62_100.0},
        {"faixa":6,"limite":4_800_000, "aliquota":0.305, "deducao":540_000.0},
    ],
}

REPARTICAO: dict[str, dict[int, dict]] = {
    "Anexo I – Comercio":     {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0276,"cofins":0.1274},3:{"pis":0.0276,"cofins":0.1274},4:{"pis":0.0276,"cofins":0.1274},5:{"pis":0.0276,"cofins":0.1274},6:{"pis":0.0276,"cofins":0.1274}},
    "Anexo II – Industria":   {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0186,"cofins":0.086}, 3:{"pis":0.0186,"cofins":0.086}, 4:{"pis":0.0186,"cofins":0.086}, 5:{"pis":0.0186,"cofins":0.086}, 6:{"pis":0.0186,"cofins":0.086}},
    "Anexo III – Servicos A": {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0167,"cofins":0.0773},3:{"pis":0.0167,"cofins":0.0773},4:{"pis":0.0167,"cofins":0.0773},5:{"pis":0.0167,"cofins":0.0773},6:{"pis":0.0167,"cofins":0.0773}},
    "Anexo IV – Servicos B":  {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0167,"cofins":0.0773},3:{"pis":0.0167,"cofins":0.0773},4:{"pis":0.0167,"cofins":0.0773},5:{"pis":0.0167,"cofins":0.0773},6:{"pis":0.0167,"cofins":0.0773}},
    "Anexo V – Servicos C":   {1:{"pis":0.0,"cofins":0.0},    2:{"pis":0.0098,"cofins":0.0454},3:{"pis":0.0098,"cofins":0.0454},4:{"pis":0.0098,"cofins":0.0454},5:{"pis":0.0098,"cofins":0.0454},6:{"pis":0.0098,"cofins":0.0454}},
}


# ─────────────────────────────────────────────────────────────
#  MODULE: utils
# ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("piscofins_pro")

_AUDIT_LOGS: list[dict] = []

def log_audit(nivel: str, cat: str, msg: str, det: str = "") -> None:
    _AUDIT_LOGS.append({
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "nivel": nivel, "categoria": cat,
        "mensagem": msg, "detalhe": det,
    })
    getattr(logger, nivel.lower() if nivel.lower() in ("info","warning","error") else "info")(
        "[{}] {} | {}".format(cat, msg, det))

def get_logs() -> list[dict]: return list(_AUDIT_LOGS)
def clear_logs() -> None: _AUDIT_LOGS.clear()

def r2(v: Optional[float]) -> Optional[float]:
    return round(v, 2) if v is not None else None

def brl(v: Optional[float]) -> str:
    if v is None: return "–"
    return "R$ {:,.2f}".format(v).replace(",","X").replace(".",",").replace("X",".")

def pct_str(v: Optional[float], casas: int = 2) -> str:
    if v is None: return "–"
    return "{:.{}f}%".format(v * 100, casas).replace(".", ",")

def fmt_mes(m: str) -> str:
    try: return datetime.strptime(m, "%Y-%m").strftime("%b/%Y")
    except: return m


# ─────────────────────────────────────────────────────────────
#  MODULE: xml_parser
# ─────────────────────────────────────────────────────────────
def _ltag(node) -> str:
    t = node.tag
    return t.split("}")[-1] if "}" in t else t

def _find(el, tag: str):
    for n in el.iter():
        if _ltag(n) == tag: return n
    return None

def _txt(el, tag: str, default: str = "") -> str:
    n = _find(el, tag)
    return n.text.strip() if n is not None and n.text else default

def _extrair_data(root) -> str:
    dh = _txt(root,"dhEmi") or _txt(root,"dEmi")
    try: return dh[:7] if dh else "SEM-DATA"
    except: return "SEM-DATA"

def _extrair_cnpj(root) -> str:
    return _txt(root,"CNPJ") or _txt(root,"CPF") or ""

@st.cache_data(show_spinner=False)
def ler_xml_bytes(nome: str, conteudo: bytes) -> tuple[list[dict], str, str]:
    """
    Retorna (itens, mes_ano, cnpj).
    Cache por (nome, hash do conteudo).
    """
    raw = conteudo.lstrip(b"\xef\xbb\xbf").strip()
    if len(raw) / 1_048_576 > UPLOAD_MAX_MB:
        raise ValueError("Arquivo excede {}MB".format(UPLOAD_MAX_MB))
    try:
        root = ET.fromstring(raw)
    except ET.ParseError as e:
        raise ValueError("XML malformado: {}".format(e))

    mes   = _extrair_data(root)
    cnpj  = _extrair_cnpj(root)
    dets  = [n for n in root.iter() if _ltag(n) == "det"]
    itens = []

    for det in dets:
        prod = _find(det, "prod")
        if not prod: continue
        ncm_raw = _txt(prod, "NCM")
        vprod   = _txt(prod, "vProd")
        cfop    = _txt(prod, "CFOP")
        cst_p   = _txt(_find(det,"PIS")    or det, "CST") if _find(det,"PIS")    else ""
        cst_c   = _txt(_find(det,"COFINS") or det, "CST") if _find(det,"COFINS") else ""
        try: valor = round(float(vprod.replace(",",".")), 2)
        except: valor = 0.0
        itens.append({
            "descricao":  _txt(prod,"xProd") or "(sem desc.)",
            "ncm_raw":    ncm_raw,
            "ncm":        ncm_raw.replace(".","").replace("-","").strip(),
            "valor":      valor,
            "cfop":       cfop,
            "cst_pis":    cst_p,
            "cst_cofins": cst_c,
        })

    log_audit("INFO","XMLParser","{} itens".format(len(itens)), nome)
    return itens, mes, cnpj


def extrair_xmls_de_zip(conteudo_zip: bytes) -> list[tuple[str, bytes]]:
    """Extrai todos os XMLs de um arquivo .zip."""
    resultado: list[tuple[str, bytes]] = []
    with zipfile.ZipFile(BytesIO(conteudo_zip)) as z:
        for nome in z.namelist():
            if nome.lower().endswith(".xml"):
                resultado.append((os.path.basename(nome), z.read(nome)))
    log_audit("INFO","ZipExtractor","{} XMLs extraidos".format(len(resultado)))
    return resultado


# ─────────────────────────────────────────────────────────────
#  MODULE: pgdas
# ─────────────────────────────────────────────────────────────
COLS_PGDAS = ["Mes","Receita_PGDAS","DAS_Pago","Receita_Monofasica_PGDAS","Segregacao_PGDAS"]

@st.cache_data(show_spinner=False)
def ler_pgdas(_file_bytes: bytes, nome: str) -> pd.DataFrame:
    """
    Le CSV/Excel do PGDAS. Colunas obrigatorias: Mes, Receita_PGDAS, DAS_Pago.
    Colunas opcionais: Receita_Monofasica_PGDAS, Segregacao_PGDAS.
    """
    if nome.lower().endswith(".csv"):
        for sep, dec in ((";",","),(","," .")):
            try:
                df = pd.read_csv(BytesIO(_file_bytes), sep=sep, decimal=dec, dtype=str)
                if df.shape[1] >= 2: break
            except: continue
    else:
        df = pd.read_excel(BytesIO(_file_bytes), dtype=str)

    df.columns = [c.strip().replace(" ","_") for c in df.columns]
    low = {c.lower(): c for c in df.columns}
    obrig = ["mes","receita_pgdas","das_pago"]
    for c in obrig:
        if c not in low:
            raise ValueError("Coluna '{}' nao encontrada. Disponiveis: {}".format(c, list(df.columns)))

    def norm_mes(v: str) -> str:
        v = str(v).strip()
        for fmt in ("%Y-%m","%m/%Y","%m-%Y","%Y/%m"):
            try: return datetime.strptime(v, fmt).strftime("%Y-%m")
            except: pass
        return v

    def to_float(v: str) -> float:
        s = str(v).strip().replace("R$","").replace(" ","")
        if "," in s and "." in s: s = s.replace(".","").replace(",",".")
        elif "," in s: s = s.replace(",",".")
        try: return round(float(s), 2)
        except: return 0.0

    rename = {low[c]: c.title().replace("_"," ") for c in obrig}
    df = df.rename(columns=rename)
    df["Mes"]           = df["Mes"].apply(norm_mes)
    df["Receita Pgdas"] = df["Receita Pgdas"].apply(to_float)
    df["Das Pago"]      = df["Das Pago"].apply(to_float)

    for opt in ["receita_monofasica_pgdas","segregacao_pgdas"]:
        col_orig = low.get(opt)
        col_novo = opt.replace("_"," ").title()
        if col_orig:
            df[col_novo] = df[col_orig].apply(to_float)
        else:
            df[col_novo] = None

    df.columns = [c.lower().replace(" ","_") for c in df.columns]
    df = df.sort_values("mes").reset_index(drop=True)
    log_audit("INFO","PGDAS","{} meses carregados".format(len(df)))
    return df


# ─────────────────────────────────────────────────────────────
#  MODULE: validators
# ─────────────────────────────────────────────────────────────
def classificar_ncm(ncm: str) -> tuple[str, str]:
    n = str(ncm).strip()
    if not n: return "INCONSISTENCIA","NCM ausente"
    if not n.isdigit(): return "INCONSISTENCIA","NCM invalido: {}".format(n)
    for t in (8,6,4):
        k = n[:t].ljust(8,"0")
        if k in TABELA_NCM: return "MONOFASICO", TABELA_NCM[k]
    return "NAO MONOFASICO","NCM fora da tabela"

def validar_tributario(ncm: str, cst_p: str, cst_c: str, cfop: str, classif: str) -> dict:
    alertas: list[str] = []
    score = 0
    is_mono = classif == "MONOFASICO"
    cf = cfop.strip().replace(".","")
    cp = cst_p.strip().lstrip("0") or "0"
    cc = cst_c.strip().lstrip("0") or "0"

    if not is_mono:
        if cp in CST_MONOFASICO or cc in CST_MONOFASICO:
            alertas.append("CST monofasico mas NCM nao consta na tabela")
            score += 40
            log_audit("RISCO","Validador","CST mono / NCM nao mono","NCM={}".format(ncm))
            return {"status_tributario":"INCONSISTENTE","score_risco":score,"motivo_alerta":" | ".join(alertas)}
        return {"status_tributario":"NAO_MONOFASICO","score_risco":0,"motivo_alerta":""}

    for label, val in (("PIS",cp),("COFINS",cc)):
        if val and val not in CST_MONOFASICO and val != "0":
            if val in CST_NORMAL:
                alertas.append("CST {} {} incompativel com monofasico".format(label,val)); score+=35
            else:
                alertas.append("CST {} {} nao reconhecido".format(label,val)); score+=15

    if cf:
        if cf in CFOP_INCOMPAT: alertas.append("CFOP {} incompativel c/ revenda monofasica".format(cfop)); score+=30
        elif cf not in CFOP_REVENDA: alertas.append("CFOP {} nao identificado como revenda padrao".format(cfop)); score+=10

    if cp and cc and cp != cc:
        alertas.append("CST PIS ({}) diverge do COFINS ({})".format(cst_p,cst_c)); score+=20

    score = min(score, 100)
    if not alertas:
        st_ = "MONOFASICO_VALIDADO"
    else:
        st_ = "MONOFASICO_COM_RISCO"
        if score >= 50: log_audit("RISCO","Validador","Score alto","NCM={} Score={}".format(ncm,score))

    return {"status_tributario":st_,"score_risco":score,
            "motivo_alerta":" | ".join(alertas) if alertas else "Validado"}


# ─────────────────────────────────────────────────────────────
#  MODULE: apuracao
# ─────────────────────────────────────────────────────────────
def agrupar_por_mes(itens: list[dict]) -> list[dict]:
    m: dict[str,dict] = {}
    for i in itens:
        k = i.get("mes","SEM-DATA")
        if k not in m: m[k] = {"tot":0.0,"mono":0.0,"mono_v":0.0,"mono_r":0.0}
        m[k]["tot"] += i["valor"]
        st = i.get("status_tributario","")
        if "MONOFASICO" in st:
            m[k]["mono"] += i["valor"]
            if st == "MONOFASICO_VALIDADO": m[k]["mono_v"] += i["valor"]
            else: m[k]["mono_r"] += i["valor"]

    out = []
    for k in sorted(x for x in m if x != "SEM-DATA") + [x for x in m if x == "SEM-DATA"]:
        t = round(m[k]["tot"],2); mo = round(m[k]["mono"],2)
        out.append({"mes":k,"receita_total":t,"receita_monofasica":mo,
                    "receita_mono_validada":round(m[k]["mono_v"],2),
                    "receita_mono_risco":round(m[k]["mono_r"],2),
                    "receita_tributavel":round(t-mo,2)})
    return out

def calcular_rbt12(agrup: list[dict], rbt12_ini: float=0.0,
                   df_pgdas: Optional[pd.DataFrame]=None) -> dict[str,float]:
    pgdas_r: dict[str,float] = {}
    if df_pgdas is not None and not df_pgdas.empty:
        pgdas_r = dict(zip(df_pgdas["mes"], df_pgdas["receita_pgdas"]))
    res:dict[str,float] = {}; hist=[]; saldo=rbt12_ini
    for row in agrup:
        mes = row["mes"]; res[mes] = round(saldo,2)
        rec = pgdas_r.get(mes, row["receita_total"])
        hist.append((mes,rec)); saldo += rec
        if len(hist) > 12: _,old = hist.pop(0); saldo -= old
    return res

def _faixa(rbt12: float, tab: list[dict]) -> Optional[dict]:
    if rbt12 <= 0: return tab[0]
    for f in tab:
        if rbt12 <= f["limite"]: return f
    return None

def _aliq_ef(rbt12: float, faixa: dict) -> float:
    if rbt12 <= 0: return 0.0
    return (rbt12 * faixa["aliquota"] - faixa["deducao"]) / rbt12

def _pct_pc(fn: int, anexo: str) -> float:
    r = REPARTICAO.get(anexo,{}).get(fn,{})
    return r.get("pis",0.0) + r.get("cofins",0.0)

def _status_credito(cb: Optional[float], cr: Optional[float]) -> dict:
    if cb is None or cr is None:
        return {"status_credito":"NAO_RECUPERAVEL","credito_final":None,"motivo_credito":"Indisponivel"}
    if cb < 0:
        log_audit("RISCO","Credito","Credito bruto negativo",brl(cb))
        return {"status_credito":"RISCO_FISCAL","credito_final":0.0,"motivo_credito":"DAS Correto > DAS Usado"}
    if cr <= 0:
        return {"status_credito":"NAO_RECUPERAVEL","credito_final":0.0,"motivo_credito":"Credito real zerado"}
    return {"status_credito":"RECUPERAVEL","credito_final":round(cr,2),"motivo_credito":""}

def _status_segregacao(mono_xml: float, mono_pgdas: Optional[float]) -> str:
    if mono_pgdas is None: return "SEM_DADO_PGDAS"
    if mono_xml <= 0: return "SEM_SEGREGACAO"
    ratio = mono_pgdas / mono_xml if mono_xml > 0 else 0.0
    if ratio < SEG_PARCIAL_MIN: return "SEM_SEGREGACAO"
    if ratio < SEG_TOTAL_MIN:   return "SEGREGACAO_PARCIAL"
    return "SEGREGACAO_TOTAL"

@st.cache_data(show_spinner=False)
def apurar_periodo_real(
    _agrup_json: str,   # JSON para compatibilidade com cache
    nome_anexo: str,
    rbt12_ini: float,
    _pgdas_json: Optional[str],
) -> list[dict]:
    """
    Pipeline de apuracao com DAS real/estimado.
    Parametros como JSON para permitir cache do Streamlit.
    """
    import json
    agrup   = json.loads(_agrup_json)
    df_pgdas = pd.read_json(io.StringIO(_pgdas_json)) if _pgdas_json else None

    tabela   = TABELAS_SIMPLES[nome_anexo]
    rbt12map = calcular_rbt12(agrup, rbt12_ini, df_pgdas)

    pgdas_idx: dict[str,dict] = {}
    if df_pgdas is not None and not df_pgdas.empty:
        for _, row in df_pgdas.iterrows():
            pgdas_idx[str(row["mes"])] = row.to_dict()

    out = []
    for row in agrup:
        mes   = row["mes"]
        rbt12 = rbt12map.get(mes,0.0)
        faixa = _faixa(rbt12,tabela)
        pr    = pgdas_idx.get(mes)

        rec_pgdas   = round(float(pr["receita_pgdas"]),2)   if pr else None
        das_real    = round(float(pr["das_pago"]),2)         if pr else None
        mono_pgdas  = round(float(pr.get("receita_monofasica_pgdas") or 0),2) if pr else None
        seg_pgdas   = pr.get("segregacao_pgdas")             if pr else None
        tem_pgdas   = pr is not None

        div = round(abs(row["receita_total"] - rec_pgdas) / rec_pgdas, 4) \
              if (rec_pgdas and rec_pgdas > 0) else None
        alerta_div = div is not None and div > DIVERGENCIA_MAX

        if alerta_div:
            log_audit("AVISO","Apuracao","Divergencia {:.1f}% em {}".format(div*100,mes),
                      "XML={} PGDAS={}".format(brl(row["receita_total"]),brl(rec_pgdas)))

        status_seg = _status_segregacao(row["receita_monofasica"], mono_pgdas)

        if faixa is None:
            out.append({**row,"rbt12":rbt12,"faixa":"ACIMA",
                        "aliquota_nominal":None,"aliquota_efetiva":None,
                        "receita_pgdas":rec_pgdas,"das_pago_real":das_real,
                        "receita_monofasica_pgdas":mono_pgdas,"segregacao_pgdas":seg_pgdas,
                        "divergencia_pct":div,"alerta_divergencia":alerta_div,
                        "tem_pgdas":tem_pgdas,"status_segregacao":status_seg,
                        "das_estimado":None,"das_usado":None,"fonte_das":"–",
                        "das_correto":None,"credito_bruto":None,
                        "pct_pis_cofins_val":None,"credito_real":None,
                        "status_credito":"NAO_RECUPERAVEL","credito_final":0.0,
                        "motivo_credito":"RBT12 acima de R$ 4,8M","alerta":"Fora do Simples"})
            log_audit("RISCO","Apuracao","RBT12 fora do Simples","Mes={}".format(mes))
            continue

        ae       = _aliq_ef(rbt12, faixa)
        das_est  = r2(row["receita_total"]      * ae)
        das_cor  = r2(row["receita_tributavel"] * ae)

        if tem_pgdas and das_real is not None:
            das_usado = r2(das_real); fonte = "REAL"
        else:
            das_usado = das_est; fonte = "ESTIMADO"

        cb   = r2((das_usado - das_cor) if das_usado is not None else None)
        ppc  = _pct_pc(faixa["faixa"], nome_anexo)
        cr   = r2(cb * ppc) if cb is not None else None
        sc   = _status_credito(cb, cr)

        out.append({**row,
            "rbt12":rbt12,"faixa":faixa["faixa"],
            "aliquota_nominal":faixa["aliquota"],"aliquota_efetiva":ae,
            "receita_pgdas":rec_pgdas,"das_pago_real":das_real,
            "receita_monofasica_pgdas":mono_pgdas,"segregacao_pgdas":seg_pgdas,
            "divergencia_pct":div,"alerta_divergencia":alerta_div,
            "tem_pgdas":tem_pgdas,"status_segregacao":status_seg,
            "das_estimado":das_est,"das_usado":das_usado,"fonte_das":fonte,
            "das_correto":das_cor,"credito_bruto":cb,
            "pct_pis_cofins_val":ppc,"credito_real":cr,**sc,"alerta":""})
    return out

def resumo_geral(itens: list[dict], aliq: float=ALIQ_ESTIMATIVA) -> dict:
    tot  = sum(i["valor"] for i in itens)
    mono = sum(i["valor"] for i in itens if "MONOFASICO" in i.get("status_tributario",""))
    nmon = sum(i["valor"] for i in itens if i.get("status_tributario")=="NAO_MONOFASICO")
    inc  = sum(i["valor"] for i in itens if i.get("status_tributario")=="INCONSISTENTE")
    mval = sum(i["valor"] for i in itens if i.get("status_tributario")=="MONOFASICO_VALIDADO")
    mris = sum(i["valor"] for i in itens if i.get("status_tributario")=="MONOFASICO_COM_RISCO")
    return {"total_geral":r2(tot),"total_monofasico":r2(mono),"total_nao_mono":r2(nmon),
            "total_inconsist":r2(inc),"total_validado":r2(mval),"total_risco":r2(mris),
            "estimativa_recup":r2(mono*aliq),"pct_monofasico":round(mono/tot,4) if tot>0 else 0.0}

def score_oportunidade(res: dict, apuracao: list[dict]) -> dict:
    s = 0
    pctm = res.get("pct_monofasico",0.0)
    if pctm >= 0.60: s+=40
    elif pctm >= 0.30: s+=25
    elif pctm >= 0.10: s+=10
    cred = sum(r.get("credito_final",0) or 0 for r in apuracao)
    if cred >= 10_000: s+=30
    elif cred >= 3_000: s+=20
    elif cred >= 500: s+=10
    n_risco = sum(1 for r in apuracao if r.get("status_credito")=="RISCO_FISCAL")
    if len(apuracao) > 0:
        if n_risco/len(apuracao) == 0: s+=20
        elif n_risco/len(apuracao) < 0.2: s+=10
    ndiv = sum(1 for r in apuracao if r.get("alerta_divergencia"))
    if ndiv == 0: s+=10
    elif ndiv <= 2: s+=5
    nivel = "ALTA" if s>=70 else ("MEDIA" if s>=40 else "BAIXA")
    emoji = {"ALTA":"🟢","MEDIA":"🟡","BAIXA":"🔴"}[nivel]
    return {"score":s,"nivel":nivel,"emoji":emoji,"credito_total":round(cred,2)}


# ─────────────────────────────────────────────────────────────
#  MODULE: auditoria
# ─────────────────────────────────────────────────────────────
SEVERITY_ORDER = {"CRITICO":0,"ALTO":1,"MEDIO":2,"BAIXO":3}

def motor_auditoria(
    itens: list[dict],
    apuracao: list[dict],
    df_pgdas: Optional[pd.DataFrame],
    res: dict,
) -> list[dict]:
    """
    Motor de auditoria tributaria.
    Retorna lista de alertas com severidade, descricao e impacto.
    """
    alertas: list[dict] = []

    def add(sev:str, mes:str, desc:str, impacto:str="", cat:str="Geral"):
        alertas.append({"severidade":sev,"mes":mes,"categoria":cat,
                        "descricao":desc,"impacto":impacto})
        log_audit("AVISO" if sev in ("CRITICO","ALTO") else "INFO",
                  "Auditoria","[{}] {}".format(sev,desc),mes)

    # 1. Divergencias XML vs PGDAS
    for r in apuracao:
        if r.get("alerta_divergencia"):
            add("ALTO", fmt_mes(r["mes"]),
                "Divergencia {:.1f}% entre XML e PGDAS".format(r["divergencia_pct"]*100),
                "Receita XML={} | PGDAS={}".format(brl(r.get("receita_total")),brl(r.get("receita_pgdas"))),
                "Divergencia")

    # 2. Credito negativo
    for r in apuracao:
        if r.get("status_credito") == "RISCO_FISCAL":
            add("CRITICO", fmt_mes(r["mes"]),
                "Credito bruto negativo – risco fiscal",
                "DAS Usado={} | DAS Correto={}".format(brl(r.get("das_usado")),brl(r.get("das_correto"))),
                "Credito Fiscal")

    # 3. Mes sem PGDAS quando PGDAS disponivel
    if df_pgdas is not None and not df_pgdas.empty:
        meses_pgdas = set(df_pgdas["mes"].tolist())
        for r in apuracao:
            if r["mes"] != "SEM-DATA" and r["mes"] not in meses_pgdas:
                add("MEDIO", fmt_mes(r["mes"]),
                    "Mes presente nos XMLs mas ausente no PGDAS",
                    "Receita XML={}".format(brl(r.get("receita_total"))), "PGDAS")

    # 4. Segregacao nao realizada
    for r in apuracao:
        if r.get("status_segregacao") == "SEM_SEGREGACAO" and r.get("receita_monofasica",0) > 0:
            add("ALTO", fmt_mes(r["mes"]),
                "Empresa nao segregou receita monofasica no PGDAS",
                "Potencial nao aproveitado: {}".format(brl(r.get("receita_monofasica"))),
                "Segregacao")
        elif r.get("status_segregacao") == "SEGREGACAO_PARCIAL":
            add("MEDIO", fmt_mes(r["mes"]),
                "Segregacao parcial detectada",
                "XML={} | PGDAS={}".format(
                    brl(r.get("receita_monofasica")), brl(r.get("receita_monofasica_pgdas"))),
                "Segregacao")

    # 5. Itens com score de risco alto
    alto_risco = [i for i in itens if i.get("score_risco",0) >= 50]
    if alto_risco:
        add("ALTO","Varios",
            "{} item(ns) com score de risco tributario >= 50".format(len(alto_risco)),
            "Valor total em risco: {}".format(brl(sum(i["valor"] for i in alto_risco))),
            "Risco Tributario")

    # 6. XML sem data de emissao
    sem_data = [i for i in itens if i.get("mes") == "SEM-DATA"]
    if sem_data:
        add("MEDIO","SEM-DATA",
            "{} item(ns) sem data de emissao (dhEmi ausente)".format(len(sem_data)),
            "Valor: {}".format(brl(sum(i["valor"] for i in sem_data))), "XML")

    # 7. DAS suspeitamente baixo
    for r in apuracao:
        das_r = r.get("das_pago_real")
        das_e = r.get("das_estimado")
        if das_r is not None and das_e is not None and das_e > 0:
            ratio = das_r / das_e
            if ratio < 0.5:
                add("MEDIO", fmt_mes(r["mes"]),
                    "DAS pago muito abaixo do estimado (ratio={:.0%})".format(ratio),
                    "DAS Real={} | DAS Estimado={}".format(brl(das_r),brl(das_e)), "DAS")

    # 8. Receita negativa/zero
    for r in apuracao:
        if r.get("receita_total",0) <= 0:
            add("ALTO", fmt_mes(r["mes"]),
                "Receita total zero ou negativa detectada","", "XML")

    alertas.sort(key=lambda x: SEVERITY_ORDER.get(x["severidade"],9))
    return alertas


# ─────────────────────────────────────────────────────────────
#  MODULE: exports
# ─────────────────────────────────────────────────────────────
def gerar_csv_apuracao(apuracao: list[dict]) -> bytes:
    rows = []
    for r in apuracao:
        rows.append({
            "Mes":            r.get("mes",""),
            "RBT12":          r.get("rbt12",""),
            "Faixa":          r.get("faixa",""),
            "Aliq. Efetiva":  "{:.2f}%".format(r["aliquota_efetiva"]*100).replace(".",",")
                              if r.get("aliquota_efetiva") is not None else "",
            "Rec. XML":       r.get("receita_total",""),
            "Rec. PGDAS":     r.get("receita_pgdas","") or "",
            "Divergencia":    "{:.2f}%".format(r["divergencia_pct"]*100).replace(".",",")
                              if r.get("divergencia_pct") is not None else "",
            "Rec. Monof.":    r.get("receita_monofasica",""),
            "Rec. Monof. PGDAS": r.get("receita_monofasica_pgdas","") or "",
            "Segregacao":     r.get("status_segregacao",""),
            "Rec. Tributavel":r.get("receita_tributavel",""),
            "DAS Usado":      r.get("das_usado",""),
            "Fonte DAS":      r.get("fonte_das",""),
            "DAS Correto":    r.get("das_correto",""),
            "Cred. Bruto":    r.get("credito_bruto",""),
            "% PIS+COF":      "{:.2f}%".format(r["pct_pis_cofins_val"]*100).replace(".",",")
                              if r.get("pct_pis_cofins_val") is not None else "",
            "Cred. Real":     r.get("credito_real",""),
            "Cred. Final":    r.get("credito_final",""),
            "Status Credito": r.get("status_credito",""),
        })
    return pd.DataFrame(rows).to_csv(index=False, sep=";", decimal=",").encode("utf-8")

def _aplicar_estilo_excel(ws, header_fill, alt_fill, money_fmt="#.##0,00", pct_f="0,00%"):
    """Aplica estilos profissionais a uma worksheet."""
    if not OPENPYXL_OK: return
    thin = Side(style="thin", color="CCCCCC")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ci, cell in enumerate(ws[1], 1):
        cell.font    = Font(bold=True, color="FFFFFF", size=10)
        cell.fill    = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border  = bdr
        ws.column_dimensions[get_column_letter(ci)].width = max(14, len(str(cell.value or ""))+4)
    ws.freeze_panes = ws["A2"]
    ws.auto_filter.ref = ws.dimensions

def gerar_excel_profissional(
    itens: list[dict], res: dict, apuracao: list[dict],
    score: dict, alertas_aud: list[dict], logs: list[dict],
) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as w:

        # ── ABA 1: Resumo Executivo ───────────────────────────
        df1 = pd.DataFrame([
            {"Indicador":"Faturamento Total",               "Valor":brl(res["total_geral"])},
            {"Indicador":"Faturamento Monofasico",          "Valor":brl(res["total_monofasico"])},
            {"Indicador":"  Validado (sem alertas)",        "Valor":brl(res["total_validado"])},
            {"Indicador":"  Com risco tributario",          "Valor":brl(res["total_risco"])},
            {"Indicador":"Faturamento Nao Monofasico",      "Valor":brl(res["total_nao_mono"])},
            {"Indicador":"% Monofasico / Total",            "Valor":pct_str(res["pct_monofasico"])},
            {"Indicador":"Credito Potencial PIS/COFINS",    "Valor":brl(score["credito_total"])},
            {"Indicador":"Score de Oportunidade",           "Valor":"{} ({}/100)".format(score["nivel"],score["score"])},
            {"Indicador":"Alertas de Auditoria",            "Valor":str(len(alertas_aud))},
        ])
        df1.to_excel(w, sheet_name="1.Resumo Executivo", index=False)

        # ── ABA 2: Apuracao Mensal ────────────────────────────
        rows2 = []
        for r in apuracao:
            rows2.append({
                "Mes":fmt_mes(r["mes"]),
                "RBT12":r.get("rbt12"),
                "Faixa":r.get("faixa","–"),
                "Aliq. Nominal":pct_str(r.get("aliquota_nominal")),
                "Aliq. Efetiva":pct_str(r.get("aliquota_efetiva"),4),
                "Rec. XML":r.get("receita_total"),
                "Rec. PGDAS":r.get("receita_pgdas"),
                "Divergencia":pct_str(r.get("divergencia_pct")) if r.get("divergencia_pct") is not None else "–",
                "Rec. Monof.":r.get("receita_monofasica"),
                "Rec. Monof. PGDAS":r.get("receita_monofasica_pgdas"),
                "Status Segregacao":r.get("status_segregacao","–"),
                "Rec. Trib.":r.get("receita_tributavel"),
                "DAS Usado":r.get("das_usado"),
                "Fonte DAS":r.get("fonte_das","–"),
                "DAS Correto":r.get("das_correto"),
                "% PIS+COF":pct_str(r.get("pct_pis_cofins_val")),
                "Cred. Bruto":r.get("credito_bruto"),
                "Cred. Final":r.get("credito_final"),
                "Status Credito":r.get("status_credito","–"),
            })
        pd.DataFrame(rows2).to_excel(w, sheet_name="2.Apuracao Mensal", index=False)

        # ── ABA 3: Itens Classificados ────────────────────────
        rows3 = [{
            "Mes":i.get("mes",""),"Arquivo":i.get("arquivo",""),
            "Descricao":i.get("descricao",""),"NCM":i.get("ncm",""),
            "CFOP":i.get("cfop",""),"CST PIS":i.get("cst_pis",""),
            "CST COFINS":i.get("cst_cofins",""),"Valor":i.get("valor"),
            "Class. NCM":i.get("classificacao",""),
            "Status Trib.":i.get("status_tributario",""),
            "Score Risco":i.get("score_risco",""),
            "Motivo Alerta":i.get("motivo_alerta",""),
        } for i in itens]
        pd.DataFrame(rows3).to_excel(w, sheet_name="3.Itens Classificados", index=False)

        # ── ABA 4: Divergencias PGDAS ─────────────────────────
        divs = [r for r in apuracao if r.get("divergencia_pct") is not None]
        if divs:
            pd.DataFrame([{
                "Mes":fmt_mes(r["mes"]),
                "Rec. XML":brl(r.get("receita_total")),
                "Rec. PGDAS":brl(r.get("receita_pgdas")),
                "Divergencia":pct_str(r.get("divergencia_pct")),
                "Alerta":("SIM" if r.get("alerta_divergencia") else "NAO"),
                "Status Seg.":r.get("status_segregacao","–"),
            } for r in divs]).to_excel(w, sheet_name="4.Divergencias PGDAS", index=False)

        # ── ABA 5: Auditoria Tributaria ───────────────────────
        if alertas_aud:
            pd.DataFrame(alertas_aud).to_excel(w, sheet_name="5.Auditoria Tributaria", index=False)
        else:
            pd.DataFrame([{"Info":"Nenhum alerta de auditoria"}]).to_excel(
                w, sheet_name="5.Auditoria Tributaria", index=False)

        # ── ABA 6: Inconsistencias NCM ────────────────────────
        inc = [i for i in itens if i.get("status_tributario") in ("INCONSISTENTE","MONOFASICO_COM_RISCO")]
        if inc:
            pd.DataFrame([{
                "Mes":i.get("mes",""),"Descricao":i.get("descricao",""),"NCM":i.get("ncm",""),
                "CFOP":i.get("cfop",""),"CST PIS":i.get("cst_pis",""),"CST COFINS":i.get("cst_cofins",""),
                "Valor":brl(i.get("valor")),"Status":i.get("status_tributario",""),
                "Score":i.get("score_risco",""),"Motivo":i.get("motivo_alerta",""),
            } for i in inc]).to_excel(w, sheet_name="6.Inconsistencias NCM", index=False)

    # Aplicar estilos post-close
    if OPENPYXL_OK:
        wb = openpyxl.load_workbook(BytesIO(output.getvalue()))
        hdr_fill = PatternFill("solid", fgColor="1E3A5F")
        alt_fill = PatternFill("solid", fgColor="EBF0F7")
        for ws in wb.worksheets:
            _aplicar_estilo_excel(ws, hdr_fill, alt_fill)
        out2 = BytesIO()
        wb.save(out2)
        return out2.getvalue()

    return output.getvalue()


def gerar_pdf(res: dict, apuracao: list[dict], score: dict, texto_exec: str) -> bytes:
    """Gera PDF executivo. Fallback silencioso se reportlab ausente."""
    if not REPORTLAB_OK:
        return b""

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    cor_azul = colors.HexColor("#1E3A5F")

    def H1(t): return Paragraph(t, ParagraphStyle("h1",parent=styles["Heading1"],
                                textColor=cor_azul,fontSize=16,spaceAfter=8))
    def H2(t): return Paragraph(t, ParagraphStyle("h2",parent=styles["Heading2"],
                                textColor=cor_azul,fontSize=12,spaceAfter=6))
    def P(t):  return Paragraph(t, styles["Normal"])

    story = [
        H1("PIS/COFINS Pro – Relatorio Executivo"),
        P("Emitido em: {}".format(datetime.now().strftime("%d/%m/%Y %H:%M"))),
        P("Versao: {}".format(VERSAO)),
        HRFlowable(width="100%", thickness=2, color=cor_azul),
        Spacer(1, 0.5*cm),
        H2("Score de Oportunidade: {} {} ({}/100)".format(score["nivel"],score["emoji"],score["score"])),
        Spacer(1,0.3*cm),
        H2("Indicadores Principais"),
    ]

    dados_tab = [
        ["Indicador","Valor"],
        ["Faturamento Total", brl(res["total_geral"])],
        ["Faturamento Monofasico", brl(res["total_monofasico"])],
        ["% Monofasico", pct_str(res["pct_monofasico"])],
        ["Credito Potencial PIS/COFINS", brl(score["credito_total"])],
        ["Monofasico Validado", brl(res["total_validado"])],
        ["Monofasico com Risco", brl(res["total_risco"])],
    ]
    t = Table(dados_tab, colWidths=[10*cm, 6*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), cor_azul),
        ("TEXTCOLOR",(0,0),(-1,0), colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#EBF0F7")]),
        ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("#CCCCCC")),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),6),
        ("BOTTOMPADDING",(0,0),(-1,-1),6),
    ]))
    story += [t, Spacer(1,0.5*cm), H2("Resumo Executivo")]

    for linha in texto_exec.split("\n"):
        if linha.strip():
            story.append(P(linha))
    story += [
        Spacer(1,0.5*cm),
        HRFlowable(width="100%", thickness=1, color=colors.HexColor("#CCCCCC")),
        P("<i>Aviso: Analise preliminar. Validar com contador habilitado. LC 123/2006.</i>"),
    ]
    doc.build(story)
    return buf.getvalue()


def texto_resumo_executivo(res: dict, apuracao: list[dict], score: dict, n_xmls: int) -> str:
    cred  = score["credito_total"]
    n_seg = sum(1 for r in apuracao if r.get("status_segregacao") == "SEGREGACAO_TOTAL")
    n_par = sum(1 for r in apuracao if r.get("status_segregacao") == "SEGREGACAO_PARCIAL")
    n_sem = sum(1 for r in apuracao if r.get("status_segregacao") == "SEM_SEGREGACAO")
    n_div = sum(1 for r in apuracao if r.get("alerta_divergencia"))
    linhas = [
        "Foram analisados {} arquivo(s) XML abrangendo {} mes(es).".format(n_xmls,len(apuracao)),
        "",
        "OPORTUNIDADE IDENTIFICADA",
        "Receita monofasica: {} ({:.1f}% do total).".format(
            brl(res["total_monofasico"]), res["pct_monofasico"]*100),
        "Credito potencial de PIS/COFINS: {}.".format(brl(cred)),
        "",
        "SITUACAO DA SEGREGACAO",
        "Meses com segregacao total: {} | Parcial: {} | Sem segregacao: {}.".format(n_seg,n_par,n_sem),
        "",
        "ALERTAS",
        "Divergencias XML vs PGDAS (>5%): {} mes(es).".format(n_div),
        "Itens com risco tributario: {}.".format(
            sum(1 for i in [] if i.get("status_tributario")=="MONOFASICO_COM_RISCO")),
        "",
        "NIVEL DE OPORTUNIDADE: {} {} ({}/100)".format(score["nivel"],score["emoji"],score["score"]),
        "",
        "Base legal: LC 123/2006 · Resolucao CGSN 140/2018 · Lei 10.147/2000",
        "Aviso: Estimativa preliminar. Recomenda-se validacao contabil.",
    ]
    return "\n".join(linhas)


# ─────────────────────────────────────────────────────────────
#  MODULE: dashboard
# ─────────────────────────────────────────────────────────────
def graficos(res: dict, apuracao: list[dict], itens: list[dict]) -> None:
    if not PLOTLY_OK:
        st.warning("Plotly nao instalado. Execute: pip install plotly")
        # Fallback nativo
        _graficos_fallback(res, apuracao)
        return

    tab1,tab2,tab3,tab4 = st.tabs(
        ["Composicao","Evolucao Mensal","Credito por Mes","Ranking NCM"])

    with tab1:
        c1,c2 = st.columns(2)
        with c1:
            fig = px.pie(
                names=["Validado","Com Risco","Nao Monofasico","Inconsistente"],
                values=[res["total_validado"],res["total_risco"],
                        res["total_nao_mono"],res["total_inconsist"]],
                title="Composicao do Faturamento",
                color_discrete_sequence=["#1a7a4a","#b8860b","#2d5a8e","#8b1a1a"],
                hole=0.4)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            vals_seg = {
                "Seg. Total":   sum(1 for r in apuracao if r.get("status_segregacao")=="SEGREGACAO_TOTAL"),
                "Seg. Parcial": sum(1 for r in apuracao if r.get("status_segregacao")=="SEGREGACAO_PARCIAL"),
                "Sem Seg.":     sum(1 for r in apuracao if r.get("status_segregacao")=="SEM_SEGREGACAO"),
            }
            fig2 = px.pie(names=list(vals_seg.keys()), values=list(vals_seg.values()),
                          title="Status de Segregacao",
                          color_discrete_sequence=["#1a7a4a","#b8860b","#8b1a1a"],hole=0.4)
            st.plotly_chart(fig2, use_container_width=True)

    with tab2:
        valids = [r for r in apuracao if r.get("receita_total") is not None]
        if valids:
            df_ev = pd.DataFrame({
                "Mes":    [fmt_mes(r["mes"]) for r in valids],
                "XML":    [r["receita_total"] for r in valids],
                "PGDAS":  [r.get("receita_pgdas") or 0 for r in valids],
                "Monof.": [r["receita_monofasica"] for r in valids],
                "Trib.":  [r["receita_tributavel"] for r in valids],
            })
            fig = px.line(df_ev, x="Mes", y=["XML","PGDAS","Monof.","Trib."],
                          title="Evolucao Mensal da Receita", markers=True)
            st.plotly_chart(fig, use_container_width=True)

    with tab3:
        valids = [r for r in apuracao if r.get("credito_final") is not None]
        if valids:
            cor_mapa = {"RECUPERAVEL":"#1a7a4a","NAO_RECUPERAVEL":"#2d5a8e","RISCO_FISCAL":"#8b1a1a"}
            fig = px.bar(
                x=[fmt_mes(r["mes"]) for r in valids],
                y=[r.get("credito_final",0) or 0 for r in valids],
                color=[r.get("status_credito","") for r in valids],
                color_discrete_map=cor_mapa,
                title="Credito Final por Mes (R$)", labels={"x":"Mes","y":"R$","color":"Status"})
            st.plotly_chart(fig, use_container_width=True)

    with tab4:
        df_ncm = (pd.DataFrame(itens)[lambda d: d["classificacao"]=="MONOFASICO"]
                  .groupby("ncm")["valor"].sum().reset_index()
                  .sort_values("valor",ascending=False).head(15))
        if not df_ncm.empty:
            fig = px.bar(df_ncm, x="valor", y="ncm", orientation="h",
                         title="Top 15 NCMs Monofasicos por Valor",
                         labels={"valor":"R$","ncm":"NCM"})
            st.plotly_chart(fig, use_container_width=True)

def _graficos_fallback(res: dict, apuracao: list[dict]) -> None:
    st.bar_chart(pd.DataFrame.from_dict({
        "Monof. Validado": res["total_validado"],
        "Monof. Risco":    res["total_risco"],
        "Nao Monofasico":  res["total_nao_mono"],
    }, orient="index", columns=["Valor (R$)"]))
    if apuracao:
        df_c = pd.DataFrame({"Mes":[fmt_mes(r["mes"]) for r in apuracao],
                             "Credito":[r.get("credito_final",0) or 0 for r in apuracao]}).set_index("Mes")
        st.bar_chart(df_c)


# ─────────────────────────────────────────────────────────────
#  MODULE: ui (pipeline principal)
# ─────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def processar_tudo(arquivos_json: str, aliq: float) -> tuple[list[dict], dict]:
    import json
    arquivos = json.loads(arquivos_json)
    todos: list[dict] = []
    vistos: set[str] = set()
    for nome, b64 in arquivos:
        import base64
        conteudo = base64.b64decode(b64)
        chave = nome + str(len(conteudo))
        if chave in vistos:
            log_audit("AVISO","Pipeline","XML duplicado ignorado",nome); continue
        vistos.add(chave)
        try:
            raw_itens, mes, cnpj = ler_xml_bytes(nome, conteudo)
        except ValueError as e:
            st.warning("Erro em '{}': {}".format(nome, e))
            log_audit("ERRO","Pipeline",str(e),nome); continue
        for i in raw_itens:
            cl, mot = classificar_ncm(i["ncm"])
            val_t   = validar_tributario(i["ncm"],i["cst_pis"],i["cst_cofins"],i["cfop"],cl)
            todos.append({**i,"arquivo":nome,"mes":mes,"cnpj":cnpj,
                          "classificacao":cl,"motivo_ncm":mot,**val_t})
    return todos, resumo_geral(todos, aliq)


def main() -> None:
    st.set_page_config(page_title="PIS/COFINS Pro",page_icon="📊",
                       layout="wide", initial_sidebar_state="expanded")
    st.markdown("""<style>
    .saas-header{background:linear-gradient(135deg,#1E3A5F,#2d5a8e);color:white;
                 padding:20px 24px;border-radius:12px;margin-bottom:16px;}
    .saas-header h1{margin:0;font-size:26px;font-weight:800;}
    .saas-header p{margin:4px 0 0;opacity:.8;font-size:13px;}
    .score-box{border-radius:14px;padding:22px;text-align:center;color:white;}
    .section-title{font-size:17px;font-weight:700;color:#1E3A5F;
                   border-left:4px solid #2d5a8e;padding-left:10px;margin:18px 0 8px;}
    div[data-testid="stMetricValue"]{font-size:20px!important;font-weight:700!important;}
    </style>""", unsafe_allow_html=True)

    st.markdown("""<div class="saas-header">
        <h1>📊 PIS/COFINS Pro &nbsp;|&nbsp; Recuperacao Tributaria – Simples Nacional</h1>
        <p>Regime de Revenda · Classificacao NCM/CST/CFOP · Apuracao PGDAS-D · v{}</p>
    </div>""".format(VERSAO), unsafe_allow_html=True)

    clear_logs()

    # ── SIDEBAR ───────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### ⚙️ Configuracoes")
        nome_anexo = st.selectbox("Anexo do Simples", list(TABELAS_SIMPLES.keys()))
        rbt12_ini  = st.number_input("RBT12 anterior (R$)", min_value=0.0,
                                     value=360_000.0, step=1_000.0, format="%.2f",
                                     help="Receita dos 12 meses anteriores ao primeiro XML.")
        usar_pgdas = st.checkbox("Usar dados reais do PGDAS", value=False)
        aliq_pct   = st.number_input("Aliquota PIS+COFINS estimativa (%)",
                                     min_value=0.0,max_value=100.0,
                                     value=9.25,step=0.05,format="%.2f")
        aliq_dec   = aliq_pct / 100
        st.markdown("---")
        st.caption("`{}` NCMs na tabela monofasica".format(len(TABELA_NCM)))
        if not PLOTLY_OK: st.warning("Plotly nao instalado.\npip install plotly")
        if not REPORTLAB_OK: st.info("PDF desativado.\npip install reportlab")

    # ── UPLOAD ────────────────────────────────────────────────
    st.markdown('<div class="section-title">Upload de Documentos Fiscais</div>',
                unsafe_allow_html=True)
    col_x, col_p = st.columns(2)

    with col_x:
        st.markdown("**XMLs de NF-e ou arquivo .zip**")
        up_xmls = st.file_uploader("Selecione XMLs ou .zip",
                                   type=["xml","zip"], accept_multiple_files=True)

    df_pgdas_raw = None
    with col_p:
        if usar_pgdas:
            st.markdown("**Planilha PGDAS/DAS**")
            with st.expander("Formato esperado"):
                st.markdown("""
| Mes | Receita_PGDAS | DAS_Pago | Receita_Monofasica_PGDAS | Segregacao_PGDAS |
|---|---|---|---|---|
| 2024-01 | 45000,00 | 1980,00 | 8000,00 | 1 |
                """)
                st.download_button("Baixar template",
                    data="Mes;Receita_PGDAS;DAS_Pago;Receita_Monofasica_PGDAS;Segregacao_PGDAS\n2024-01;45000,00;1980,00;8000,00;1\n".encode(),
                    file_name="template_pgdas_v7.csv",mime="text/csv")
            up_pgdas = st.file_uploader("CSV ou Excel do PGDAS",
                                        type=["csv","xlsx","xls"],key="pgdas")
            if up_pgdas:
                try:
                    df_pgdas_raw = ler_pgdas(up_pgdas.read(), up_pgdas.name)
                    st.success("{} mes(es): {} a {}".format(
                        len(df_pgdas_raw),fmt_mes(df_pgdas_raw["mes"].min()),
                        fmt_mes(df_pgdas_raw["mes"].max())))
                except ValueError as e:
                    st.error(str(e))
        else:
            st.info("Ative 'Usar dados reais do PGDAS' na sidebar.")

    if not up_xmls:
        st.info("📂 Aguardando upload dos XMLs para iniciar a analise.")
        st.stop()

    # ── COLETA DE ARQUIVOS (inclui zip) ───────────────────────
    import base64, json
    arquivos_raw: list[tuple[str,bytes]] = []
    for f in up_xmls:
        conteudo = f.read()
        if f.name.lower().endswith(".zip"):
            try: arquivos_raw.extend(extrair_xmls_de_zip(conteudo))
            except Exception as e: st.warning("Erro no zip '{}': {}".format(f.name,e))
        else:
            arquivos_raw.append((f.name, conteudo))

    if not arquivos_raw:
        st.error("Nenhum XML encontrado nos arquivos enviados."); st.stop()

    # Serializa para cache
    arquivos_json = json.dumps([(n, base64.b64encode(b).decode()) for n,b in arquivos_raw])
    pgdas_json    = df_pgdas_raw.to_json() if df_pgdas_raw is not None else None

    # ── PROCESSAMENTO ─────────────────────────────────────────
    prog = st.progress(0, "Lendo e classificando XMLs...")
    itens, res = processar_tudo(arquivos_json, aliq_dec)
    prog.progress(40, "Agrupando por mes...")

    if not itens:
        st.error("Nenhum item extraido."); st.stop()

    agrup = agrupar_por_mes(itens)
    prog.progress(60, "Apurando PGDAS-D...")

    agrup_json = json.dumps(agrup)
    apuracao   = apurar_periodo_real(agrup_json, nome_anexo, rbt12_ini, pgdas_json)
    prog.progress(80, "Gerando analises...")

    score      = score_oportunidade(res, apuracao)
    alertas_a  = motor_auditoria(itens, apuracao, df_pgdas_raw, res)
    texto_exec = texto_resumo_executivo(res, apuracao, score, len(arquivos_raw))
    df_itens   = pd.DataFrame(itens)
    prog.progress(100, "Concluido!"); prog.empty()

    # ── NIVEL 1 vs NIVEL 2 ────────────────────────────────────
    nivel = "2 – Apuracao Real" if usar_pgdas and df_pgdas_raw is not None else "1 – Triagem"
    st.info("🔎 Modo de Analise: **Nivel {}**".format(nivel), icon="📋")

    # ── SCORE ─────────────────────────────────────────────────
    st.markdown('<div class="section-title">Score de Oportunidade</div>',unsafe_allow_html=True)
    _,sc,_ = st.columns([2,3,2])
    with sc:
        cor = {"ALTA":"#1a7a4a","MEDIA":"#b8860b","BAIXA":"#8b1a1a"}[score["nivel"]]
        st.markdown("""<div class="score-box" style="background:{}">
            <div style="font-size:48px">{}</div>
            <div style="font-size:30px;font-weight:800">{} OPORTUNIDADE</div>
            <div style="font-size:18px;opacity:.9">Score: {}/100</div>
            <div style="font-size:16px;margin-top:8px;opacity:.85">Credito Potencial: {}</div>
        </div>""".format(cor,score["emoji"],score["nivel"],score["score"],
                         brl(score["credito_total"])), unsafe_allow_html=True)

    # ── METRICAS ──────────────────────────────────────────────
    st.markdown('<div class="section-title">Indicadores Executivos</div>',unsafe_allow_html=True)
    m1,m2,m3,m4,m5,m6 = st.columns(6)
    m1.metric("Total Analisado",    brl(res["total_geral"]))
    m2.metric("Monofasico",         brl(res["total_monofasico"]), pct_str(res["pct_monofasico"]))
    m3.metric("Validado",           brl(res["total_validado"]))
    m4.metric("Com Risco",          brl(res["total_risco"]))
    m5.metric("Credito Potencial",  brl(score["credito_total"]))
    m6.metric("Alertas Auditoria",  str(len(alertas_a)))

    # ── AUDITORIA ─────────────────────────────────────────────
    if alertas_a:
        st.markdown('<div class="section-title">Auditoria Tributaria</div>',unsafe_allow_html=True)
        sev_icon = {"CRITICO":"🔴","ALTO":"🟠","MEDIO":"🟡","BAIXO":"🔵"}
        for al in alertas_a:
            icon = sev_icon.get(al["severidade"],"⚪")
            msg  = "{} **[{}]** {} – {} | {}".format(
                icon, al["severidade"], al["mes"], al["descricao"], al["impacto"])
            if al["severidade"] in ("CRITICO","ALTO"): st.error(msg)
            elif al["severidade"] == "MEDIO": st.warning(msg)
            else: st.info(msg)

    # ── DASHBOARD ─────────────────────────────────────────────
    st.markdown('<div class="section-title">Dashboard Executivo</div>',unsafe_allow_html=True)
    graficos(res, apuracao, itens)

    # ── APURACAO MENSAL ───────────────────────────────────────
    st.markdown('<div class="section-title">Apuracao PGDAS-D – Credito por Mes</div>',unsafe_allow_html=True)
    linhas_ap = []
    seg_icons = {"SEGREGACAO_TOTAL":"✅","SEGREGACAO_PARCIAL":"🟡",
                 "SEM_SEGREGACAO":"🔴","SEM_DADO_PGDAS":"–"}
    for r in apuracao:
        l = {"Mes":fmt_mes(r["mes"]),"RBT12":brl(r.get("rbt12")),
             "Faixa":str(r.get("faixa","–")),"Aliq. Ef.":pct_str(r.get("aliquota_efetiva"),4),
             "Rec. XML":brl(r.get("receita_total"))}
        if usar_pgdas:
            l["Rec. PGDAS"]  = brl(r.get("receita_pgdas"))
            l["Divergencia"] = ("🔴 " if r.get("alerta_divergencia") else "✅ ") + \
                               pct_str(r.get("divergencia_pct")) if r.get("divergencia_pct") is not None else "–"
            l["Segregacao"]  = seg_icons.get(r.get("status_segregacao","–"),"–") + " " + \
                               (r.get("status_segregacao","–") or "–")
        l.update({"Rec. Monof.":brl(r.get("receita_monofasica")),
                  "Rec. Trib.": brl(r.get("receita_tributavel")),
                  "DAS Usado":  brl(r.get("das_usado")),
                  "Fonte":      ("🟢 REAL" if r.get("fonte_das")=="REAL" else "🟡 EST."),
                  "DAS Correto":brl(r.get("das_correto")),
                  "% PIS+COF":  pct_str(r.get("pct_pis_cofins_val")),
                  "Cred. Final":brl(r.get("credito_final")),
                  "Status":     {"RECUPERAVEL":"✅","NAO_RECUPERAVEL":"⚪","RISCO_FISCAL":"🔴"}
                                .get(r.get("status_credito",""),"–")})
        linhas_ap.append(l)
    st.dataframe(pd.DataFrame(linhas_ap), use_container_width=True, hide_index=True)

    valid = [r for r in apuracao if r.get("credito_final") is not None]
    t1,t2,t3,t4 = st.columns(4)
    t1.metric("DAS Total Usado",     brl(sum(r.get("das_usado",0) or 0 for r in valid)))
    t2.metric("DAS Total Correto",   brl(sum(r.get("das_correto",0) or 0 for r in valid)))
    t3.metric("Credito Bruto Total", brl(sum(r.get("credito_bruto",0) or 0 for r in valid)))
    t4.metric("Credito Real PIS/COF",brl(score["credito_total"]),
              delta="a recuperar" if score["credito_total"]>0 else None)

    # ── OPORTUNIDADE vs RISCOS ────────────────────────────────
    oc,rc = st.columns(2)
    with oc:
        st.markdown('<div class="section-title">Oportunidade Identificada</div>',unsafe_allow_html=True)
        top3 = sorted([r for r in apuracao if (r.get("credito_final") or 0)>0],
                      key=lambda x:x.get("credito_final",0),reverse=True)[:3]
        for r in top3:
            st.success("**{}** – Credito: {} | Monof.: {}".format(
                fmt_mes(r["mes"]),brl(r.get("credito_final")),brl(r.get("receita_monofasica"))))
        if not top3: st.info("Nenhum credito recuperavel.")
    with rc:
        st.markdown('<div class="section-title">Riscos Identificados</div>',unsafe_allow_html=True)
        riscos = [i for i in itens if i.get("status_tributario")=="MONOFASICO_COM_RISCO"]
        for i in riscos[:4]:
            st.warning("**{}** | NCM {} | Score {} | {}".format(
                i["descricao"][:30],i["ncm"],i.get("score_risco",0),
                i.get("motivo_alerta","")[:50]))
        if not riscos: st.success("Nenhum risco tributario.")

    # ── ITENS CLASSIFICADOS ───────────────────────────────────
    st.markdown('<div class="section-title">Itens Classificados</div>',unsafe_allow_html=True)
    f1,f2,f3 = st.columns(3)
    filtro_st  = f1.selectbox("Status",["Todos","MONOFASICO_VALIDADO","MONOFASICO_COM_RISCO",
                                        "NAO_MONOFASICO","INCONSISTENTE"])
    filtro_mes = f2.selectbox("Mes",["Todos"]+sorted(df_itens["mes"].unique().tolist()))
    filtro_arq = f3.selectbox("Arquivo",["Todos"]+sorted(df_itens["arquivo"].unique().tolist()))
    df_ex = df_itens.copy()
    if filtro_st  != "Todos": df_ex = df_ex[df_ex["status_tributario"]==filtro_st]
    if filtro_mes != "Todos": df_ex = df_ex[df_ex["mes"]==filtro_mes]
    if filtro_arq != "Todos": df_ex = df_ex[df_ex["arquivo"]==filtro_arq]
    st.dataframe(df_ex[["mes","arquivo","descricao","ncm","cfop","cst_pis","cst_cofins",
                          "valor","status_tributario","score_risco","motivo_alerta"]].rename(columns={
        "mes":"Mes","arquivo":"Arquivo","descricao":"Descricao","ncm":"NCM","cfop":"CFOP",
        "cst_pis":"CST PIS","cst_cofins":"CST COF","valor":"Valor (R$)",
        "status_tributario":"Status Trib.","score_risco":"Score","motivo_alerta":"Motivo",
    }), use_container_width=True, height=360)

    # ── RESUMO EXECUTIVO ──────────────────────────────────────
    with st.expander("Resumo Executivo Automatico", expanded=False):
        st.text(texto_exec)

    # ── EXPORTS ───────────────────────────────────────────────
    st.markdown('<div class="section-title">Exportar Relatorios</div>',unsafe_allow_html=True)
    e1,e2,e3 = st.columns(3)
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    with e1:
        st.download_button("📥 Excel Profissional (6 abas)",
            data=gerar_excel_profissional(itens,res,apuracao,score,alertas_a,get_logs()),
            file_name="piscofins_pro_{}.xlsx".format(ts),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with e2:
        st.download_button("📥 Apuracao CSV",
            data=gerar_csv_apuracao(apuracao),
            file_name="apuracao_{}.csv".format(ts),
            mime="text/csv", use_container_width=True)
    with e3:
        if REPORTLAB_OK:
            pdf = gerar_pdf(res, apuracao, score, texto_exec)
            st.download_button("📄 Relatorio PDF",
                data=pdf, file_name="relatorio_{}.pdf".format(ts),
                mime="application/pdf", use_container_width=True)
        else:
            st.info("PDF: pip install reportlab")

    # ── LOGS ──────────────────────────────────────────────────
    with st.expander("Logs de Auditoria ({} eventos)".format(len(get_logs()))):
        df_log = pd.DataFrame(get_logs())
        if not df_log.empty:
            icons = {"INFO":"🔵","AVISO":"🟡","ERRO":"🔴","RISCO":"🟠"}
            df_log["nivel"] = df_log["nivel"].apply(lambda x: icons.get(x,"")+x)
            st.dataframe(df_log, use_container_width=True, hide_index=True)

    # ── METODOLOGIA ───────────────────────────────────────────
    with st.expander("Metodologia e Base Legal"):
        st.markdown("""
**Nivel 1 – Triagem:** apenas XMLs. Classificacao por NCM + validacao CST/CFOP.

**Nivel 2 – Apuracao Real:** XML + PGDAS. DAS real substitui estimado. Segregacao validada.

**Status de Segregacao:**
- SEGREGACAO_TOTAL: PGDAS >= 90% do monofasico XML
- SEGREGACAO_PARCIAL: PGDAS entre 5% e 90%
- SEM_SEGREGACAO: PGDAS < 5% (credito potencial maximo)

**Aliquota Efetiva PGDAS-D:**
`ae = (RBT12 × aliquota_nominal − deducao) / RBT12`

**Credito:**
`Cred. Bruto = DAS Usado − DAS Correto`
`Cred. Real  = Cred. Bruto × (% PIS + % COFINS)`

**Base legal:** LC 123/2006 · CGSN 140/2018 · Lei 10.147/2000 · Art. 18 LC 123
        """)

    st.divider()
    st.caption("PIS/COFINS Pro v{} · Analise preliminar · Validar com contador · LC 123/2006".format(VERSAO))


if __name__ == "__main__":
    main()

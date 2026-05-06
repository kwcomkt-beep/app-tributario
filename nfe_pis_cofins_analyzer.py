"""
=============================================================
  ANALISADOR PIS/COFINS MONOFASICO + APURACAO PGDAS-D v4
  Simples Nacional · Regime de Revenda · CST 04
  RBT12 dinamico · Faixa por mes · Credito real PIS/COFINS
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
#     Cada faixa: {faixa, limite, aliquota, deducao}
# ─────────────────────────────────────────────────────────────

ANEXO_I = [
    {"faixa": 1, "limite": 180_000,     "aliquota": 0.04,  "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,     "aliquota": 0.073, "deducao": 5_940.0},
    {"faixa": 3, "limite": 720_000,     "aliquota": 0.095, "deducao": 13_860.0},
    {"faixa": 4, "limite": 1_800_000,   "aliquota": 0.107, "deducao": 22_500.0},
    {"faixa": 5, "limite": 3_600_000,   "aliquota": 0.143, "deducao": 87_300.0},
    {"faixa": 6, "limite": 4_800_000,   "aliquota": 0.19,  "deducao": 378_000.0},
]

ANEXO_II = [
    {"faixa": 1, "limite": 180_000,     "aliquota": 0.045, "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,     "aliquota": 0.078, "deducao": 5_940.0},
    {"faixa": 3, "limite": 720_000,     "aliquota": 0.10,  "deducao": 13_860.0},
    {"faixa": 4, "limite": 1_800_000,   "aliquota": 0.113, "deducao": 22_500.0},
    {"faixa": 5, "limite": 3_600_000,   "aliquota": 0.147, "deducao": 85_500.0},
    {"faixa": 6, "limite": 4_800_000,   "aliquota": 0.30,  "deducao": 720_000.0},
]

ANEXO_III = [
    {"faixa": 1, "limite": 180_000,     "aliquota": 0.06,  "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,     "aliquota": 0.112, "deducao": 9_360.0},
    {"faixa": 3, "limite": 720_000,     "aliquota": 0.135, "deducao": 17_640.0},
    {"faixa": 4, "limite": 1_800_000,   "aliquota": 0.16,  "deducao": 35_640.0},
    {"faixa": 5, "limite": 3_600_000,   "aliquota": 0.21,  "deducao": 125_640.0},
    {"faixa": 6, "limite": 4_800_000,   "aliquota": 0.33,  "deducao": 648_000.0},
]

ANEXO_IV = [
    {"faixa": 1, "limite": 180_000,     "aliquota": 0.045, "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,     "aliquota": 0.09,  "deducao": 8_100.0},
    {"faixa": 3, "limite": 720_000,     "aliquota": 0.102, "deducao": 12_420.0},
    {"faixa": 4, "limite": 1_800_000,   "aliquota": 0.14,  "deducao": 39_780.0},
    {"faixa": 5, "limite": 3_600_000,   "aliquota": 0.22,  "deducao": 183_780.0},
    {"faixa": 6, "limite": 4_800_000,   "aliquota": 0.33,  "deducao": 828_000.0},
]

ANEXO_V = [
    {"faixa": 1, "limite": 180_000,     "aliquota": 0.15,  "deducao": 0.0},
    {"faixa": 2, "limite": 360_000,     "aliquota": 0.18,  "deducao": 5_400.0},
    {"faixa": 3, "limite": 720_000,     "aliquota": 0.195, "deducao": 13_500.0},
    {"faixa": 4, "limite": 1_800_000,   "aliquota": 0.205, "deducao": 20_700.0},
    {"faixa": 5, "limite": 3_600_000,   "aliquota": 0.23,  "deducao": 62_100.0},
    {"faixa": 6, "limite": 4_800_000,   "aliquota": 0.305, "deducao": 540_000.0},
]

TABELAS_SIMPLES = {
    "Anexo I – Comercio":     ANEXO_I,
    "Anexo II – Industria":   ANEXO_II,
    "Anexo III – Servicos A": ANEXO_III,
    "Anexo IV – Servicos B":  ANEXO_IV,
    "Anexo V – Servicos C":   ANEXO_V,
}

# ─────────────────────────────────────────────────────────────
#  3. TABELA DE REPARTICAO DO DAS – LC 123/2006
#     Percentuais de PIS e COFINS sobre a aliquota efetiva
#     por faixa e por anexo.
#     Faixa 1 do Anexo I: PIS/COFINS = 0 (isencao na 1a faixa)
#     Demais faixas: conforme resolucao CGSN 140/2018
# ─────────────────────────────────────────────────────────────

REPARTICAO = {
    "Anexo I – Comercio": {
        1: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0,    "pis": 0.0,    "cpp": 0.0,    "icms": 1.0},
        2: {"cpf": 0.0, "csll": 0.0, "cofins": 0.1274, "pis": 0.0276, "cpp": 0.4368, "icms": 0.4082},
        3: {"cpf": 0.0, "csll": 0.0, "cofins": 0.1274, "pis": 0.0276, "cpp": 0.4368, "icms": 0.4082},
        4: {"cpf": 0.0, "csll": 0.0, "cofins": 0.1274, "pis": 0.0276, "cpp": 0.4368, "icms": 0.4082},
        5: {"cpf": 0.0, "csll": 0.0, "cofins": 0.1274, "pis": 0.0276, "cpp": 0.4368, "icms": 0.4082},
        6: {"cpf": 0.0, "csll": 0.0, "cofins": 0.1274, "pis": 0.0276, "cpp": 0.4368, "icms": 0.4082},
    },
    "Anexo II – Industria": {
        1: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0,    "pis": 0.0,    "cpp": 0.0,    "icms": 0.0,   "ipi": 1.0},
        2: {"cpf": 0.0, "csll": 0.0, "cofins": 0.086,  "pis": 0.0186, "cpp": 0.4054, "icms": 0.3768, "ipi": 0.1132},
        3: {"cpf": 0.0, "csll": 0.0, "cofins": 0.086,  "pis": 0.0186, "cpp": 0.4054, "icms": 0.3768, "ipi": 0.1132},
        4: {"cpf": 0.0, "csll": 0.0, "cofins": 0.086,  "pis": 0.0186, "cpp": 0.4054, "icms": 0.3768, "ipi": 0.1132},
        5: {"cpf": 0.0, "csll": 0.0, "cofins": 0.086,  "pis": 0.0186, "cpp": 0.4054, "icms": 0.3768, "ipi": 0.1132},
        6: {"cpf": 0.0, "csll": 0.0, "cofins": 0.086,  "pis": 0.0186, "cpp": 0.4054, "icms": 0.3768, "ipi": 0.1132},
    },
    "Anexo III – Servicos A": {
        1: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0,    "pis": 0.0,    "cpp": 0.0,    "iss": 1.0},
        2: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "cpp": 0.2816, "iss": 0.3333, "irpj": 0.04, "csll2": 0.0411},
        3: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "cpp": 0.2816, "iss": 0.3333, "irpj": 0.04, "csll2": 0.0411},
        4: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "cpp": 0.2816, "iss": 0.3333, "irpj": 0.04, "csll2": 0.0411},
        5: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "cpp": 0.2816, "iss": 0.3333, "irpj": 0.04, "csll2": 0.0411},
        6: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "cpp": 0.2816, "iss": 0.3333, "irpj": 0.04, "csll2": 0.0411},
    },
    "Anexo IV – Servicos B": {
        1: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0,    "pis": 0.0,    "iss": 1.0},
        2: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "iss": 0.4444, "irpj": 0.18, "csll2": 0.1516},
        3: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "iss": 0.4444, "irpj": 0.18, "csll2": 0.1516},
        4: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "iss": 0.4444, "irpj": 0.18, "csll2": 0.1516},
        5: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "iss": 0.4444, "irpj": 0.18, "csll2": 0.1516},
        6: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0773, "pis": 0.0167, "iss": 0.4444, "irpj": 0.18, "csll2": 0.1516},
    },
    "Anexo V – Servicos C": {
        1: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0,    "pis": 0.0,    "iss": 1.0},
        2: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0454, "pis": 0.0098, "iss": 0.2,   "irpj": 0.2718, "csll2": 0.073, "cpp": 0.4},
        3: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0454, "pis": 0.0098, "iss": 0.2,   "irpj": 0.2718, "csll2": 0.073, "cpp": 0.4},
        4: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0454, "pis": 0.0098, "iss": 0.2,   "irpj": 0.2718, "csll2": 0.073, "cpp": 0.4},
        5: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0454, "pis": 0.0098, "iss": 0.2,   "irpj": 0.2718, "csll2": 0.073, "cpp": 0.4},
        6: {"cpf": 0.0, "csll": 0.0, "cofins": 0.0454, "pis": 0.0098, "iss": 0.2,   "irpj": 0.2718, "csll2": 0.073, "cpp": 0.4},
    },
}

ALIQUOTA_PIS_COFINS_ESTIMATIVA = 0.0925


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
    """Extrai dhEmi/dEmi do XML e retorna 'YYYY-MM'."""
    dh = text_local(root, "dhEmi") or text_local(root, "dEmi")
    if not dh:
        return "SEM-DATA"
    try:
        return dh[:7]
    except Exception:
        return "SEM-DATA"


def ler_xml_nfe(conteudo):
    """
    Le bytes de NF-e XML.
    Retorna (lista_itens, mes_ano_str).
    Robusto: ignora namespace, remove BOM.
    """
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
#  5. CLASSIFICACAO NCM
# ─────────────────────────────────────────────────────────────
def classificar_item(ncm, tabela):
    """
    Retorna (classificacao, motivo).
    Busca por 8, 6 ou 4 digitos iniciais.
    """
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
#  6. AGRUPAMENTO MENSAL
# ─────────────────────────────────────────────────────────────
def agrupar_por_mes(itens_classificados):
    """
    Agrupa itens classificados por mes (YYYY-MM).
    Retorna lista ordenada cronologicamente de dicts:
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
    # Ordena: meses validos cronologicamente, SEM-DATA no final
    chaves_validas  = sorted([k for k in meses if k != "SEM-DATA"])
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
#  7. RBT12 DINAMICO
# ─────────────────────────────────────────────────────────────
def calcular_rbt12(agrupamento_mensal, rbt12_inicial=0.0):
    """
    Calcula o RBT12 rolling para cada mes.

    Logica:
      - Para o primeiro mes: usa rbt12_inicial (informado pelo usuario)
        como proxy dos 12 meses anteriores ao periodo dos XMLs.
      - Para meses subsequentes: desloca a janela somando o mes atual
        e subtraindo o mais antigo (quando ha 12+ meses disponíveis).

    Retorna dict {mes: rbt12}.

    Nota: para calculo rigoroso do PGDAS-D o ideal e ter o historico
    completo de 12 meses. O rbt12_inicial supre a ausencia desse historico.
    """
    meses = [r["mes"] for r in agrupamento_mensal]
    receitas = {r["mes"]: r["receita_total"] for r in agrupamento_mensal}

    # Janela deslizante de ate 12 meses
    resultado = {}
    historico = []  # lista de (mes, receita) dos ultimos 12

    # Inicializa com rbt12_inicial como "bloco" representando os 12 meses anteriores
    saldo = rbt12_inicial

    for mes in meses:
        # RBT12 deste mes = soma dos 12 meses ANTERIORES a ele
        resultado[mes] = round(saldo, 2)

        # Adiciona este mes ao historico e atualiza saldo
        historico.append((mes, receitas[mes]))
        saldo += receitas[mes]

        # Se historico > 12 meses, remove o mais antigo
        if len(historico) > 12:
            _, rec_antiga = historico.pop(0)
            saldo -= rec_antiga

    return resultado


# ─────────────────────────────────────────────────────────────
#  8. IDENTIFICACAO DE FAIXA
# ─────────────────────────────────────────────────────────────
def identificar_faixa(rbt12, tabela_anexo):
    """
    Recebe o RBT12 e a lista de faixas do anexo.
    Retorna o dict da faixa correspondente ou None se acima do limite.
    """
    if rbt12 <= 0:
        return tabela_anexo[0]  # Faixa 1 como padrao para RBT12 zero
    for faixa in tabela_anexo:
        if rbt12 <= faixa["limite"]:
            return faixa
    return None  # Acima de R$ 4,8M – fora do Simples


# ─────────────────────────────────────────────────────────────
#  9. CALCULO DA ALIQUOTA EFETIVA (PGDAS-D)
# ─────────────────────────────────────────────────────────────
def calcular_aliquota_efetiva(rbt12, faixa):
    """
    Formula oficial PGDAS-D:
      aliquota_efetiva = (RBT12 * aliquota_nominal - deducao) / RBT12

    Retorna float da aliquota efetiva, ou 0.0 se RBT12 = 0.
    """
    if rbt12 <= 0:
        return 0.0
    return (rbt12 * faixa["aliquota"] - faixa["deducao"]) / rbt12


# ─────────────────────────────────────────────────────────────
#  10. PERCENTUAL PIS/COFINS DA REPARTICAO
# ─────────────────────────────────────────────────────────────
def calcular_pct_pis_cofins(faixa_num, nome_anexo):
    """
    Retorna o percentual combinado PIS + COFINS para a faixa
    informada, conforme tabela de reparticao do DAS.
    """
    rep = REPARTICAO.get(nome_anexo, {})
    faixa_rep = rep.get(faixa_num, {})
    return faixa_rep.get("pis", 0.0) + faixa_rep.get("cofins", 0.0)


# ─────────────────────────────────────────────────────────────
#  11. PIPELINE DE APURACAO MES A MES
# ─────────────────────────────────────────────────────────────
def apurar_periodo(agrupamento_mensal, nome_anexo, rbt12_inicial=0.0):
    """
    Pipeline completo de apuracao do credito mensal.

    Para cada mes calcula:
      - RBT12 dinamico
      - Faixa do Simples
      - Aliquota efetiva (formula PGDAS-D)
      - DAS pago (sobre receita total)
      - DAS correto (sobre receita tributavel)
      - Credito bruto (diferenca dos DAS)
      - Percentual PIS+COFINS da reparticao
      - Credito real PIS/COFINS

    Retorna lista de dicts com todos os campos.
    """
    tabela_anexo = TABELAS_SIMPLES[nome_anexo]
    rbt12_por_mes = calcular_rbt12(agrupamento_mensal, rbt12_inicial)

    resultado = []
    for row in agrupamento_mensal:
        mes   = row["mes"]
        rbt12 = rbt12_por_mes.get(mes, 0.0)

        faixa = identificar_faixa(rbt12, tabela_anexo)

        if faixa is None:
            # RBT12 acima do limite do Simples
            resultado.append({
                **row,
                "rbt12":           rbt12,
                "faixa":           "ACIMA",
                "aliquota_efetiva":None,
                "das_pago":        None,
                "das_correto":     None,
                "credito_bruto":   None,
                "pct_pis_cofins":  None,
                "credito_real":    None,
                "alerta":          "RBT12 acima de R$ 4,8M – fora do Simples",
            })
            continue

        aliq_ef  = calcular_aliquota_efetiva(rbt12, faixa)
        das_pago    = row["receita_total"]     * aliq_ef
        das_correto = row["receita_tributavel"] * aliq_ef
        cred_bruto  = das_pago - das_correto
        pct_pc      = calcular_pct_pis_cofins(faixa["faixa"], nome_anexo)
        cred_real   = cred_bruto * pct_pc

        resultado.append({
            **row,
            "rbt12":           rbt12,
            "faixa":           faixa["faixa"],
            "aliquota_nominal":faixa["aliquota"],
            "aliquota_efetiva":aliq_ef,
            "das_pago":        das_pago,
            "das_correto":     das_correto,
            "credito_bruto":   cred_bruto,
            "pct_pis_cofins":  pct_pc,
            "credito_real":    cred_real,
            "alerta":          "",
        })

    return resultado


# ─────────────────────────────────────────────────────────────
#  12. RESUMO GERAL (mantido)
# ─────────────────────────────────────────────────────────────
def calcular_resumo(itens_classificados, aliquota=ALIQUOTA_PIS_COFINS_ESTIMATIVA):
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
#  13. PIPELINE PRINCIPAL
# ─────────────────────────────────────────────────────────────
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
#  14. EXPORTACOES
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


def gerar_csv_apuracao(apuracao):
    linhas = []
    for r in apuracao:
        linhas.append({
            "Mes":              r["mes"],
            "RBT12":            r["rbt12"],
            "Faixa":            r["faixa"],
            "Aliq. Efetiva":    r["aliquota_efetiva"] if r["aliquota_efetiva"] is not None else "",
            "Rec. Total":       r["receita_total"],
            "Rec. Monofasica":  r["receita_monofasica"],
            "Rec. Tributavel":  r["receita_tributavel"],
            "DAS Pago":         r["das_pago"]      if r["das_pago"]      is not None else "",
            "DAS Correto":      r["das_correto"]   if r["das_correto"]   is not None else "",
            "Credito Bruto":    r["credito_bruto"] if r["credito_bruto"] is not None else "",
            "% PIS+COFINS":     r["pct_pis_cofins"] if r["pct_pis_cofins"] is not None else "",
            "Credito Real":     r["credito_real"]  if r["credito_real"]  is not None else "",
        })
    df = pd.DataFrame(linhas)
    return df.to_csv(index=False, sep=";", decimal=",").encode("utf-8")


# ─────────────────────────────────────────────────────────────
#  15. HELPERS DE FORMATACAO
# ─────────────────────────────────────────────────────────────
def brl(v):
    if v is None:
        return "–"
    return "R$ {:,.2f}".format(v).replace(",","X").replace(".",",").replace("X",".")

def pct(v, casas=4):
    if v is None:
        return "–"
    fmt = "{{:.{}f}}%".format(casas)
    return fmt.format(v * 100)

def fmt_mes(m):
    try:
        return datetime.strptime(m, "%Y-%m").strftime("%b/%Y")
    except Exception:
        return m


# ─────────────────────────────────────────────────────────────
#  16. INTERFACE STREAMLIT
# ─────────────────────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="PIS/COFINS + PGDAS-D",
        page_icon="📊",
        layout="wide",
    )

    st.title("📊 Analisador PIS/COFINS Monofasico + Apuracao PGDAS-D")
    st.caption(
        "Simples Nacional · Regime de Revenda · "
        "RBT12 dinamico · Faixa mensal · Credito real PIS/COFINS"
    )
    st.divider()

    # ── SIDEBAR ───────────────────────────────────────────────
    with st.sidebar:
        st.header("Configuracoes")

        st.subheader("Simples Nacional")
        nome_anexo = st.selectbox(
            "Anexo",
            list(TABELAS_SIMPLES.keys()),
            help="Selecione o anexo da sua atividade"
        )

        rbt12_inicial = st.number_input(
            "RBT12 anterior ao periodo (R$)",
            min_value=0.0,
            value=360_000.0,
            step=1_000.0,
            format="%.2f",
            help=(
                "Informe a receita bruta acumulada dos 12 meses "
                "anteriores ao primeiro mes dos XMLs carregados. "
                "Usado como base para o calculo dinamico do RBT12."
            )
        )

        st.markdown("---")
        st.subheader("PIS/COFINS (estimativa simples)")
        aliquota_pc = st.number_input(
            "Aliquota estimativa (%)",
            min_value=0.0, max_value=100.0,
            value=9.25, step=0.05, format="%.2f",
        )
        aliquota_decimal = aliquota_pc / 100

        st.markdown("---")
        st.markdown(
            "**Tabela NCM:**  \n"
            "`{}` NCMs monofasicos  \n"
            "_Edite `TABELA_NCM_MONOFASICO` para atualizar._".format(
                len(TABELA_NCM_MONOFASICO)
            )
        )

        # Preview reparticao do anexo selecionado
        st.markdown("---")
        with st.expander("Ver reparticao do anexo"):
            rep = REPARTICAO.get(nome_anexo, {})
            rows_rep = []
            for faixa_n, tributos in rep.items():
                pis_v    = tributos.get("pis", 0.0)
                cofins_v = tributos.get("cofins", 0.0)
                rows_rep.append({
                    "Faixa": faixa_n,
                    "PIS":    pct(pis_v, 2),
                    "COFINS": pct(cofins_v, 2),
                    "PIS+COFINS": pct(pis_v + cofins_v, 2),
                })
            st.dataframe(pd.DataFrame(rows_rep), hide_index=True, use_container_width=True)

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
    c1.metric("Faturamento Total",           brl(resumo["total_geral"]))
    c2.metric("Faturamento Monofasico",      brl(resumo["total_monofasico"]))
    c3.metric("Faturamento Nao Monofasico",  brl(resumo["total_nao_mono"]))
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

    # ── APURACAO PGDAS-D AVANCADA ─────────────────────────────
    st.subheader("4  Apuracao PGDAS-D Avancada – Credito Real PIS/COFINS")

    agrupado  = agrupar_por_mes(itens)
    apuracao  = apurar_periodo(agrupado, nome_anexo, rbt12_inicial)

    # Alertas de fora do Simples
    alertas = [r for r in apuracao if r.get("alerta")]
    for al in alertas:
        st.warning("Mes {}: {}".format(fmt_mes(al["mes"]), al["alerta"]))

    # Monta tabela exibicao
    linhas = []
    for r in apuracao:
        linhas.append({
            "Mes":            fmt_mes(r["mes"]),
            "RBT12":          brl(r["rbt12"]),
            "Faixa":          r["faixa"],
            "Aliq. Nominal":  pct(r.get("aliquota_nominal"), 2),
            "Aliq. Efetiva":  pct(r.get("aliquota_efetiva"), 4),
            "Rec. Total":     brl(r["receita_total"]),
            "Rec. Monof.":    brl(r["receita_monofasica"]),
            "Rec. Tributavel":brl(r["receita_tributavel"]),
            "DAS Pago":       brl(r.get("das_pago")),
            "DAS Correto":    brl(r.get("das_correto")),
            "Cred. Bruto":    brl(r.get("credito_bruto")),
            "% PIS+COF":      pct(r.get("pct_pis_cofins"), 2),
            "Cred. Real":     brl(r.get("credito_real")),
        })

    st.dataframe(pd.DataFrame(linhas), use_container_width=True, hide_index=True)

    # Totais
    st.markdown("---")
    valid = [r for r in apuracao if r.get("credito_real") is not None]
    total_cred_bruto = sum(r["credito_bruto"] for r in valid)
    total_cred_real  = sum(r["credito_real"]  for r in valid)
    total_das_pago   = sum(r["das_pago"]      for r in valid)
    total_das_corr   = sum(r["das_correto"]   for r in valid)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total DAS Pago",       brl(total_das_pago))
    m2.metric("Total DAS Correto",    brl(total_das_corr))
    m3.metric("Credito Bruto Total",  brl(total_cred_bruto))
    m4.metric(
        "Credito Real PIS/COFINS",
        brl(total_cred_real),
        delta="a recuperar" if total_cred_real > 0 else None
    )

    if total_cred_real <= 0 and valid:
        st.info("Nenhum credito real apurado. Verifique os XMLs, o RBT12 anterior e o anexo selecionado.")

    # Nota metodologica
    with st.expander("Metodologia de calculo"):
        st.markdown("""
**RBT12 dinamico:**
Para o primeiro mes, o RBT12 e igual ao valor informado como *RBT12 anterior ao periodo*.
Para os meses seguintes, a janela deslizante de 12 meses e atualizada automaticamente.

**Aliquota efetiva (PGDAS-D):**
```
aliquota_efetiva = (RBT12 × aliquota_nominal − deducao) / RBT12
```

**DAS pago vs. DAS correto:**
```
DAS_pago    = Receita Total     × aliquota_efetiva
DAS_correto = Rec. Tributavel   × aliquota_efetiva
Credito bruto = DAS_pago − DAS_correto
```

**Credito real PIS/COFINS:**
```
Credito_real = Credito_bruto × (% PIS + % COFINS da reparticao)
```

A tabela de reparticao segue a LC 123/2006 e Resolucao CGSN 140/2018.
Na faixa 1 do Anexo I, PIS e COFINS sao zero (isencao).
        """)

    # Export CSV
    csv_bytes = gerar_csv_apuracao(apuracao)
    st.download_button(
        label="Baixar apuracao CSV",
        data=csv_bytes,
        file_name="apuracao_pgdas_avancado.csv",
        mime="text/csv",
    )

    # ── ALERTAS INCONSISTENCIAS ───────────────────────────────
    inconsist = df[df["classificacao"] == "INCONSISTENCIA"]
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

    sem_data = df[df["mes"] == "SEM-DATA"]
    if not sem_data.empty:
        st.warning(
            "{} item(ns) sem data de emissao (dhEmi ausente). "
            "Agrupados como 'SEM-DATA'.".format(len(sem_data))
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

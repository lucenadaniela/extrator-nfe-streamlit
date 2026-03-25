from __future__ import annotations

import io
import os
import re
import json
import time
import unicodedata
import xml.etree.ElementTree as ET
from io import BytesIO
from typing import Optional, Dict, Any, List, Tuple

import pandas as pd
import streamlit as st

# --- geocoding opcional (PLUMA) ---
try:
    from geopy.geocoders import Nominatim
    _GEOPY_OK = True
except Exception:
    _GEOPY_OK = False

import folium
from streamlit_folium import st_folium


# =========================================================
# CONFIG GERAL
# =========================================================
st.set_page_config(page_title="Extrator NF-e (Clientes)", page_icon="📄", layout="wide")

st.markdown("""
<style>
.main .block-container {max-width: 1200px;}
.stDownloadButton > button {border-radius: 10px; padding: 10px 16px; font-weight: 600;}
.kpi {text-align:center; border-radius:14px; padding:16px; border:1px solid #ececf3; background:#fff;}
.kpi .label {color:#4b5563; font-size:12px;}
.kpi .value {font-weight:800; font-size:24px; color:#111827;}
.leaflet-control-attribution {font-size:12px !important;}
</style>
""", unsafe_allow_html=True)

CLIENTES = [
    "PLUMA ESPUMAS LTDA (TXT)",
    "NORSA REFRIGERANTES S.A (XML)",
]

st.title("📄 Extrator NF-e → Excel (por cliente)")
st.caption("Escolha o cliente e envie o arquivo no formato correto. Eu gero a planilha certinha por layout.")


# =========================================================
# =====================  PLUMA (TXT)  =====================
# =========================================================

# -------- Config IBGE / Fretes --------
GEOCACHE_FILE = "geocache_destinos.json"
PLOT_HEIGHT = 520

CAPITAL_IBGE_PE = {"RECIFE"}
RMR_IBGE_PE = {
    "RECIFE", "OLINDA", "JABOATAO DOS GUARARAPES", "PAULISTA",
    "CABO DE SANTO AGOSTINHO", "IPOJUCA", "CAMARAGIBE",
    "SAO LOURENCO DA MATA", "ABREU E LIMA", "IGARASSU",
    "ITAPISSUMA", "ARACOIABA", "MORENO", "ILHA DE ITAMARACA",
    "FERNANDO DE NORONHA",
}

VALOR_FRETE_RECIFE_CAPITAL  = 130.00
VALOR_FRETE_RECIFE_METRO    = 130.00
VALOR_FRETE_RECIFE_INTERIOR = 249.00

# Interior PE por região (baseado na lista de municípios por RD)
VALOR_FRETE_PE_MATA_NORTE = 249.00
VALOR_FRETE_PE_MATA_SUL = 249.00
VALOR_FRETE_PE_AGRESTE_CENTRAL = 249.00
VALOR_FRETE_PE_AGRESTE_MERIDIONAL = 249.00
VALOR_FRETE_PE_AGRESTE_SETENTRIONAL = 249.00
VALOR_FRETE_PE_SERTAO_MOXOTO = 360.00
VALOR_FRETE_PE_SERTAO_PAJEU = 380.00
VALOR_FRETE_PE_SERTAO_ITAPARICA = 400.00
VALOR_FRETE_PE_SERTAO_CENTRAL = 440.00
VALOR_FRETE_PE_SERTAO_ARARIPE = 470.00
VALOR_FRETE_PE_SERTAO_SAO_FRANCISCO = 490.00

PE_REGIOES_FRETE = {
    "MATA NORTE": VALOR_FRETE_PE_MATA_NORTE,
    "MATA SUL": VALOR_FRETE_PE_MATA_SUL,
    "AGRESTE CENTRAL": VALOR_FRETE_PE_AGRESTE_CENTRAL,
    "AGRESTE MERIDIONAL": VALOR_FRETE_PE_AGRESTE_MERIDIONAL,
    "AGRESTE SETENTRIONAL": VALOR_FRETE_PE_AGRESTE_SETENTRIONAL,
    "SERTAO DO MOXOTO": VALOR_FRETE_PE_SERTAO_MOXOTO,
    "SERTAO DO PAJEU": VALOR_FRETE_PE_SERTAO_PAJEU,
    "SERTAO DE ITAPARICA": VALOR_FRETE_PE_SERTAO_ITAPARICA,
    "SERTAO CENTRAL": VALOR_FRETE_PE_SERTAO_CENTRAL,
    "SERTAO DO ARARIPE": VALOR_FRETE_PE_SERTAO_ARARIPE,
    "SERTAO DO SAO FRANCISCO": VALOR_FRETE_PE_SERTAO_SAO_FRANCISCO,
}

PE_REGIOES_IBGE = {
    "MATA NORTE": {
        "ALIANCA", "BUENOS AIRES", "CAMUTANGA", "CARPINA", "CHA DE ALEGRIA",
        "CONDADO", "FERREIROS", "GLORIA DO GOITA", "GOIANA", "ITAMBE",
        "ITAQUITINGA", "LAGOA DO CARRO", "LAGOA DO ITAENGA", "MACAPARANA",
        "NAZARE DA MATA", "PAUDALHO", "TIMBAUBA", "TRACUNHAEM", "VICENCIA",
    },
    "MATA SUL": {
        "AGUA PRETA", "AMARAJI", "BARREIROS", "BELEM DE MARIA", "CATENDE",
        "CHA GRANDE", "CORTES", "ESCADA", "GAMELEIRA", "JAQUEIRA",
        "JOAQUIM NABUCO", "MARAIAL", "PALMARES", "POMBOS", "PRIMAVERA",
        "QUIPAPA", "RIBEIRAO", "RIO FORMOSO", "SAO BENEDITO DO SUL",
        "SAO JOSE DA COROA GRANDE", "SIRINHAEM", "TAMANDARE",
        "VITORIA DE SANTO ANTAO", "XEXEU",
    },
    "AGRESTE CENTRAL": {
        "AGRESTINA", "ALAGOINHA", "ALTINHO", "BARRA DE GUABIRABA",
        "BELO JARDIM", "BEZERROS", "BONITO", "BREJO DA MADRE DE DEUS",
        "CACHOEIRINHA", "CAMOCIM DE SAO FELIX", "CARUARU", "CUPIRA",
        "GRAVATA", "IBIRAJUBA", "JATAUBA", "LAGOA DOS GATOS", "PANELAS",
        "PESQUEIRA", "POCAO", "RIACHO DAS ALMAS", "SAIRE", "SANHARO",
        "SAO BENTO DO UNA", "SAO CAETANO", "SAO JOAQUIM DO MONTE", "TACAIMBO",
    },
    "AGRESTE MERIDIONAL": {
        "AGUAS BELAS", "ANGELIM", "BOM CONSELHO", "BREJAO", "BUIQUE",
        "CAETES", "CALCADO", "CANHOTINHO", "CAPOEIRAS", "CORRENTES",
        "GARANHUNS", "IATI", "ITAIBA", "JUCATI", "JUPI", "JUREMA",
        "LAGOA DO OURO", "LAJEDO", "PALMEIRINA", "PARANATAMA", "PEDRA",
        "SALOA", "SAO JOAO", "TEREZINHA", "TUPANATINGA", "VENTUROSA",
    },
    "AGRESTE SETENTRIONAL": {
        "BOM JARDIM", "CASINHAS", "CUMARU", "FEIRA NOVA", "FREI MIGUELINHO",
        "JOAO ALFREDO", "LIMOEIRO", "MACHADOS", "OROBO", "PASSIRA",
        "SALGADINHO", "SANTA CRUZ DO CAPIBARIBE", "SANTA MARIA DO CAMBUCA",
        "SAO VICENTE FERRER", "SURUBIM", "TAQUARITINGA DO NORTE", "TORITAMA",
        "VERTENTE DO LERIO", "VERTENTES",
    },
    "SERTAO DO ARARIPE": {
        "ARARIPINA", "BODOCO", "EXU", "GRANITO", "IPUBI", "MOREILANDIA",
        "OURICURI", "SANTA CRUZ", "SANTA FILOMENA", "TRINDADE",
    },
    "SERTAO CENTRAL": {
        "CEDRO", "MIRANDIBA", "PARNAMIRIM", "SALGUEIRO",
        "SAO JOSE DO BELMONTE", "SERRITA", "TERRA NOVA", "VERDEJANTE",
    },
    "SERTAO DE ITAPARICA": {
        "BELEM DE SAO FRANCISCO", "CARNAUBEIRA DA PENHA", "FLORESTA",
        "ITACURUBA", "JATOBA", "PETROLANDIA", "TACARATU",
    },
    "SERTAO DO MOXOTO": {
        "ARCOVERDE", "BETANIA", "CUSTODIA", "IBIMIRIM", "INAJA", "MANARI", "SERTANIA",
    },
    "SERTAO DO PAJEU": {
        "AFOGADOS DA INGAZEIRA", "BREJINHO", "CALUMBI", "CARNAIBA", "FLORES",
        "IGUARACI", "INGAZEIRA", "ITAPETIM", "QUIXABA", "SANTA CRUZ DA BAIXA VERDE",
        "SANTA TEREZINHA", "SAO JOSE DO EGITO", "SERRA TALHADA", "SOLIDAO",
        "TABIRA", "TRIUNFO", "TUPARETAMA",
    },
    "SERTAO DO SAO FRANCISCO": {
        "AFRANIO", "CABROBO", "DORMENTES", "LAGOA GRANDE", "OROCO",
        "PETROLINA", "SANTA MARIA DA BOA VISTA",
    },
}

PE_MUNICIPIO_ALIASES = {
    "CABO DO STO. AGOSTINHO": "CABO DE SANTO AGOSTINHO",
    "CABO DO STO AGOSTINHO": "CABO DE SANTO AGOSTINHO",
    "CABO DE STO AGOSTINHO": "CABO DE SANTO AGOSTINHO",
    "JABOATAO DOSGUARARAPES": "JABOATAO DOS GUARARAPES",
    "ILHA DE ITAMARACA": "ILHA DE ITAMARACA",
    "ITAMARACA": "ILHA DE ITAMARACA",
    "BELEM DE S. FRANCISCO": "BELEM DE SAO FRANCISCO",
    "BELEM DE S FRANCISCO": "BELEM DE SAO FRANCISCO",
    "CARNAUBEIRA DA PENHA": "CARNAUBEIRA DA PENHA",
    "CAMAUBEIRA DA PENHA": "CARNAUBEIRA DA PENHA",
    "STA. CRUZ DA BAIXA VERDE": "SANTA CRUZ DA BAIXA VERDE",
    "STA CRUZ DA BAIXA VERDE": "SANTA CRUZ DA BAIXA VERDE",
    "STA. MARIA DO CAMBUCA": "SANTA MARIA DO CAMBUCA",
    "STA MARIA DO CAMBUCA": "SANTA MARIA DO CAMBUCA",
    "STA. CRUZ DO CAPIBARIBE": "SANTA CRUZ DO CAPIBARIBE",
    "STA CRUZ DO CAPIBARIBE": "SANTA CRUZ DO CAPIBARIBE",
    "TUPARATEMA": "TUPARETAMA",
    "PARNAMIRIM": "PARNAMIRIM",
}

CAPITAL_IBGE_AL = {"MACEIO"}
RMM_MACEIO_IBGE = {
    "ATALAIA", "BARRA DE SANTO ANTONIO", "BARRA DE SAO MIGUEL", "COQUEIRO SECO",
    "MACEIO", "MARECHAL DEODORO", "MESSIAS", "MURICI", "PARIPUEIRA", "PILAR",
    "RIO LARGO", "SANTA LUZIA DO NORTE", "SATUBA",
}

VALOR_FRETE_MACEIO_CAPITAL  = 100.00
VALOR_FRETE_MACEIO_METRO    = 200.00
VALOR_FRETE_MACEIO_INTERIOR = 200.00


def normalizar_municipio_pe(municipio: str) -> str:
    key = strip_accents_upper(sanitize_municipio_name(municipio))
    return PE_MUNICIPIO_ALIASES.get(key, key)


def obter_regiao_pe(municipio: str) -> Optional[str]:
    key = normalizar_municipio_pe(municipio)
    for regiao, municipios in PE_REGIOES_IBGE.items():
        if key in municipios:
            return regiao
    return None
NUM_BR = r'(\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?)'
TEL_PAT = re.compile(r'(\(?\d{2}\)?\s*\d{4,5}[-\s]?\d{4}|\b\d{10,11}\b)')
RE_CEP_EXATO = re.compile(r'\b\d{5}-\d{3}\b')


def strip_accents_upper(s: str) -> str:
    if not s:
        return ""
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.upper().strip()


def sanitize_municipio_name(raw: str) -> str:
    if not raw:
        return ""
    s = str(raw).strip()
    s = s.split(",")[0]
    s = re.sub(r"\b(PE|PERNAMBUCO|AL|ALAGOAS)\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s


def detectar_origem_por_municipios(registros) -> str:
    municipios = []
    for reg in registros:
        mun = reg.get("MUNICÍPIO")
        if not mun:
            continue
        mun_clean = sanitize_municipio_name(mun)
        if mun_clean:
            municipios.append(strip_accents_upper(mun_clean))

    if not municipios:
        return "Recife (PE)"

    pe_hits = sum(1 for k in municipios if k in CAPITAL_IBGE_PE or k in RMR_IBGE_PE)
    al_hits = sum(1 for k in municipios if k in CAPITAL_IBGE_AL or k in RMM_MACEIO_IBGE)

    if al_hits > pe_hits:
        return "Maceió (AL)"
    if al_hits > 0 and pe_hits == 0:
        return "Maceió (AL)"
    return "Recife (PE)"


def classificar_zona_ibge(municipio: str, origem: str):
    key = normalizar_municipio_pe(municipio)
    origem_norm = strip_accents_upper(origem)

    if "RECIFE" in origem_norm:
        if key in CAPITAL_IBGE_PE:
            return "Capital", VALOR_FRETE_RECIFE_CAPITAL
        if key in RMR_IBGE_PE:
            return "Metropolitana", VALOR_FRETE_RECIFE_METRO

        regiao_pe = obter_regiao_pe(key)
        if regiao_pe:
            return regiao_pe, PE_REGIOES_FRETE[regiao_pe]

        return "Interior", VALOR_FRETE_RECIFE_INTERIOR

    if "MACEIO" in origem_norm:
        if key in CAPITAL_IBGE_AL:
            return "Capital", VALOR_FRETE_MACEIO_CAPITAL
        if key in RMM_MACEIO_IBGE:
            return "Metropolitana", VALOR_FRETE_MACEIO_METRO
        return "Interior", VALOR_FRETE_MACEIO_INTERIOR

    # fallback
    if key in CAPITAL_IBGE_AL:
        return "Capital", VALOR_FRETE_MACEIO_CAPITAL
    if key in RMM_MACEIO_IBGE:
        return "Metropolitana", VALOR_FRETE_MACEIO_METRO
    if key in CAPITAL_IBGE_PE:
        return "Capital", VALOR_FRETE_RECIFE_CAPITAL
    if key in RMR_IBGE_PE:
        return "Metropolitana", VALOR_FRETE_RECIFE_METRO

    regiao_pe = obter_regiao_pe(key)
    if regiao_pe:
        return regiao_pe, PE_REGIOES_FRETE[regiao_pe]

    return "Interior", VALOR_FRETE_RECIFE_INTERIOR

def prox_nao_vazia(linhas, j, max_look=15):
    n = len(linhas)
    k = j
    passos = 0
    while k < n and passos < max_look:
        val = (linhas[k] or "").strip()
        if val:
            return val
        k += 1
        passos += 1
    return ""


def normalize_phone(raw: str) -> str:
    if not raw:
        return ""
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 11:
        return f"({digits[:2]}){digits[2:7]}-{digits[7:]}"
    if len(digits) == 10:
        return f"({digits[:2]}){digits[2:6]}-{digits[6:]}"
    return re.sub(r"\s+", " ", raw).strip()


def find_phone_in_text(txt: str) -> str:
    m = TEL_PAT.search(txt or "")
    return normalize_phone(m.group(1)) if m else ""


def extrair_cep_exato(txt: str) -> str:
    if not txt:
        return ""
    m = RE_CEP_EXATO.search(txt)
    return m.group(0) if m else ""


def extrair_notas_de_texto(texto: str):
    linhas = [l.strip() for l in texto.splitlines()]
    registros = []
    atual = None
    n = len(linhas)
    i = 0

    while i < n:
        linha = linhas[i]

        m_num = re.search(r'\bN[ºo]\s*([0-9\.\-]+)', linha, flags=re.IGNORECASE)
        if m_num:
            if atual:
                registros.append(atual)
            atual = {
                "Nº": m_num.group(1),
                "NOME / RAZÃO SOCIAL": None,
                "ENDEREÇO": None,
                "BAIRRO / DISTRITO": None,
                "MUNICÍPIO": None,
                "CEP": None,
                "VALOR TOTAL DA NOTA": None,
                "QUANTIDADE": None,
                "PESO BRUTO": None,
                "TELEFONE / FAX": None,
                "TELEFONE 2": None,
                "VENDEDOR": None,
                "INTEGRACAO": None,
                "CUBAGEM": None,
            }
            i += 1
            continue

        if atual is not None:
            if re.fullmatch(r'NOME\s*/\s*RAZÃO SOCIAL', linha, flags=re.IGNORECASE) and atual["NOME / RAZÃO SOCIAL"] is None:
                v = prox_nao_vazia(linhas, i + 1)
                if v:
                    atual["NOME / RAZÃO SOCIAL"] = v

            if re.fullmatch(r'ENDEREÇO', linha, flags=re.IGNORECASE):
                v = prox_nao_vazia(linhas, i + 1)
                if v:
                    atual["ENDEREÇO"] = v

            if re.fullmatch(r'BAIRRO\s*/\s*DISTRITO', linha, flags=re.IGNORECASE):
                v = prox_nao_vazia(linhas, i + 1)
                if v:
                    atual["BAIRRO / DISTRITO"] = v

            if re.fullmatch(r'MUNICÍPIO', linha, flags=re.IGNORECASE):
                v = prox_nao_vazia(linhas, i + 1)
                if v:
                    v_limpo = re.split(r'\bUF\b|CEP|\bPE\b|\bAL\b|\d{2}:\d{2}:\d{2}', v, maxsplit=1)[0].strip(" -")
                    atual["MUNICÍPIO"] = v_limpo or v

            if re.fullmatch(r'VALOR TOTAL DA NOTA', linha, flags=re.IGNORECASE):
                v = prox_nao_vazia(linhas, i + 1)
                if v:
                    m_val = re.search(NUM_BR, v)
                    atual["VALOR TOTAL DA NOTA"] = (m_val.group(1) if m_val else v).replace("R$", "").strip()

            # PESO BRUTO
            if re.fullmatch(r'PESO\s+BRUTO', linha, flags=re.IGNORECASE) and not atual["PESO BRUTO"]:
                v = prox_nao_vazia(linhas, i + 1, max_look=10)
                if v:
                    m = re.search(NUM_BR, v)
                    if m:
                        atual["PESO BRUTO"] = m.group(1)
            if not atual["PESO BRUTO"]:
                m_inline = re.search(r'\bPESO\s+BRUTO\b.*?' + NUM_BR, linha, flags=re.IGNORECASE)
                if m_inline:
                    atual["PESO BRUTO"] = m_inline.group(1)
            if not atual["PESO BRUTO"] and ("VOLUME" in linha.upper() or "NUMERAÇÃO" in linha.upper()):
                m = re.search(NUM_BR + r'(?!.*\d)', linha)
                if m:
                    atual["PESO BRUTO"] = m.group(1)

            # QUANTIDADE
            if atual["QUANTIDADE"] is None:
                m_qt = re.search(r'(\d+)\s+VOLUMES', linha, flags=re.IGNORECASE)
                if m_qt:
                    atual["QUANTIDADE"] = m_qt.group(1)
                else:
                    if re.search(r'\bQUANTIDADE\b', linha, flags=re.IGNORECASE):
                        v = prox_nao_vazia(linhas, i + 1, max_look=5)
                        m_next = re.search(r'(\d+)\s+VOLUMES', v, flags=re.IGNORECASE) or re.search(r'\b(\d+)\b', v)
                        if m_next:
                            atual["QUANTIDADE"] = m_next.group(1)

            # CEP
            if linha.strip().upper() == "CEP" and not atual["CEP"]:
                v = prox_nao_vazia(linhas, i + 1, max_look=15)
                atual["CEP"] = extrair_cep_exato(v)
            elif not atual["CEP"]:
                atual["CEP"] = extrair_cep_exato(linha)

            # TELEFONE / FAX
            if re.fullmatch(r'TELEFONE\s*/\s*FAX', linha, flags=re.IGNORECASE) and not atual["TELEFONE / FAX"]:
                v = prox_nao_vazia(linhas, i + 1, max_look=6)
                if not re.match(r'INSCRIÇÃO|DESTINAT[ÁA]RIO', v, flags=re.IGNORECASE):
                    tel = find_phone_in_text(v) or find_phone_in_text(linha)
                    if tel:
                        atual["TELEFONE / FAX"] = tel
            if not atual["TELEFONE / FAX"] and "TELEFONE / FAX" in linha.upper():
                if not re.search(r'INSCRIÇÃO\s+ESTADUAL', linha, flags=re.IGNORECASE):
                    tel = find_phone_in_text(linha)
                    if tel:
                        atual["TELEFONE / FAX"] = tel

            # TELEFONE 2
            if ("RASTREAMENTO" in linha.upper()) or ("ENTREGAID" in linha.upper()) or ("TELEFONE 2" in linha.upper()):
                m_t2 = re.search(r'TELEFONE\s*2\s*:\s*([^;]*)', linha, flags=re.IGNORECASE)
                if m_t2:
                    valor = m_t2.group(1).strip()
                    if valor:
                        atual["TELEFONE 2"] = find_phone_in_text(valor)

            # VENDEDOR / INTEGRACAO
            if "#VENDEDOR" in linha.upper() and not atual.get("VENDEDOR"):
                m_vend = re.search(r'#VENDEDOR\s*:\s*(.*?)\s+NOSSO\s+PEDIDO', linha, flags=re.IGNORECASE)
                if m_vend:
                    atual["VENDEDOR"] = m_vend.group(1).strip(" :;-")

            if "INTEGRACAO" in linha.upper() and not atual.get("INTEGRACAO"):
                prox = prox_nao_vazia(linhas, i + 1, max_look=3)
                bloco = linha + " " + (prox or "")
                m_int = re.search(r'Integracao\s*:\s*(.+?)(?:-+\s*EntregaID\b|;|\Z)', bloco, flags=re.IGNORECASE)
                if m_int:
                    atual["INTEGRACAO"] = m_int.group(1).strip(" :;-")

            # CUBAGEM
            if "CUBAGEM" in linha.upper() and not atual.get("CUBAGEM"):
                prox = prox_nao_vazia(linhas, i + 1, max_look=2)
                bloco = linha + " " + (prox or "")
                m_cub = re.search(r'CUBAGEM\s*[:\-]\s*([^;]+)', bloco, flags=re.IGNORECASE)
                if m_cub:
                    atual["CUBAGEM"] = m_cub.group(1).strip()

        i += 1

    if atual:
        registros.append(atual)
    return registros


def salvar_excel_bytes_pluma(registros, origem: str) -> Tuple[bytes, pd.DataFrame]:
    df = pd.DataFrame(registros)

    # une linhas quebradas por mesmo Nº consecutivo
    df_clean = []
    skip_next = False
    for i in range(len(df)):
        if skip_next:
            skip_next = False
            continue
        row = df.iloc[i]
        if i + 1 < len(df) and str(df.iloc[i + 1]["Nº"]) == str(row["Nº"]):
            combined = df.iloc[i + 1].copy()
            combined["Nº"] = row["Nº"]
            for campo in [
                "PESO BRUTO", "QUANTIDADE",
                "TELEFONE / FAX", "TELEFONE 2",
                "VENDEDOR", "INTEGRACAO",
                "NOME / RAZÃO SOCIAL", "ENDEREÇO",
                "BAIRRO / DISTRITO", "MUNICÍPIO", "CEP",
                "VALOR TOTAL DA NOTA", "CUBAGEM",
            ]:
                if not str(combined.get(campo, "")).strip():
                    combined[campo] = row.get(campo)
            df_clean.append(combined)
            skip_next = True
        else:
            df_clean.append(row)
    df = pd.DataFrame(df_clean).reset_index(drop=True)

    # classifica zona e frete
    zonas, fretes = [], []
    for mun in df["MUNICÍPIO"].fillna(""):
        z, f = classificar_zona_ibge(mun, origem)
        zonas.append(z)
        fretes.append(f)
    df["ZONA"] = zonas
    df["VALOR FRETE"] = fretes

    colunas = [
        "Nº", "NOME / RAZÃO SOCIAL", "ENDEREÇO",
        "BAIRRO / DISTRITO", "MUNICÍPIO", "CEP",
        "VALOR TOTAL DA NOTA", "QUANTIDADE", "PESO BRUTO",
        "CUBAGEM",
        "TELEFONE / FAX", "TELEFONE 2",
        "VENDEDOR", "INTEGRACAO",
        "ZONA", "VALOR FRETE",
    ]
    df = df.reindex(columns=colunas)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Notas")
    buffer.seek(0)
    return buffer.read(), df


# -------- Geocache / mapa (PLUMA) --------
def load_geocache() -> dict:
    if os.path.exists(GEOCACHE_FILE):
        try:
            with open(GEOCACHE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_geocache(cache: dict):
    try:
        with open(GEOCACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


COORDS_FALLBACK_RAW = {
    "RECIFE, PE": (-8.0476, -34.8770),
    "OLINDA, PE": (-8.0101, -34.8545),
    "JABOATAO DOS GUARARAPES, PE": (-8.1120, -35.0140),
    "PAULISTA, PE": (-7.9400, -34.8731),
    "CABO DE SANTO AGOSTINHO, PE": (-8.2822, -35.0320),
    "IPOJUCA, PE": (-8.3983, -35.0639),
    "CAMARAGIBE, PE": (-8.0207, -34.9786),
    "SAO LOURENCO DA MATA, PE": (-8.0062, -35.0199),
    "ABREU E LIMA, PE": (-7.9007, -34.9027),
    "IGARASSU, PE": (-7.8340, -34.9069),
    "ITAPISSUMA, PE": (-7.7750, -34.8954),
    "ARACOIABA, PE": (-7.7913, -35.0800),
    "MORENO, PE": (-8.1180, -35.0920),
    "ILHA DE ITAMARACA, PE": (-7.7478, -34.8332),
    "ILHA DE ITAMARACÁ, PE": (-7.7478, -34.8332),
}


def _norm_place_key(s: str) -> str:
    return strip_accents_upper(s)


COORDS_FALLBACK_NORM = {_norm_place_key(k): v for k, v in COORDS_FALLBACK_RAW.items()}


def geocode_city(city: str, uf: str) -> Optional[Tuple[float, float]]:
    city_clean = sanitize_municipio_name(city)
    key_raw = f"{city_clean}, {uf}"
    key_norm = _norm_place_key(key_raw)

    cache = st.session_state.get("geocache", {})

    if key_raw in cache and cache[key_raw]:
        return tuple(cache[key_raw])

    if key_norm in COORDS_FALLBACK_NORM:
        latlon = COORDS_FALLBACK_NORM[key_norm]
        cache[key_raw] = latlon
        st.session_state["geocache"] = cache
        save_geocache(cache)
        return latlon

    if _GEOPY_OK:
        try:
            estado_nome = "Alagoas" if uf == "AL" else "Pernambuco"
            geolocator = Nominatim(user_agent="nf_extractor_ws")
            loc = geolocator.geocode(f"{city_clean}, {estado_nome}, Brazil", timeout=10)
            if loc:
                latlon = (loc.latitude, loc.longitude)
                cache[key_raw] = latlon
                st.session_state["geocache"] = cache
                save_geocache(cache)
                time.sleep(1.0)
                return latlon
        except Exception:
            pass

    cache[key_raw] = None
    st.session_state["geocache"] = cache
    save_geocache(cache)
    return None


def _color_for(z):
    if z == "Capital":
        return "red"
    if z == "Metropolitana":
        return "blue"
    return "gray"


def build_map_folium(df_destinos: pd.DataFrame, origem: str):
    if df_destinos.empty:
        st.info("Sem destinos válidos para plotar no mapa.")
        return

    origem_norm = strip_accents_upper(origem)
    center = [-9.6658, -35.7353] if "MACEIO" in origem_norm else [-8.0476, -34.8770]
    m = folium.Map(location=center, zoom_start=7, tiles="OpenStreetMap")

    lats, lons = [], []
    for _, row in df_destinos.iterrows():
        lat = row["lat"]
        lon = row["lon"]
        qtd = int(row.get("QTD_ENTREGAS", 1))

        lats.append(lat)
        lons.append(lon)

        cor = _color_for(row["ZONA"])

        html_marker = f"""
        <div style="
            background:{cor};
            color:white;
            border:2px solid white;
            border-radius:50%;
            width:30px;
            height:30px;
            line-height:26px;
            text-align:center;
            font-size:12px;
            font-weight:700;
            box-shadow:0 0 4px rgba(0,0,0,0.35);
        ">
            {qtd}
        </div>
        """

        folium.Marker(
            location=[lat, lon],
            icon=folium.DivIcon(html=html_marker),
            popup=folium.Popup(
                html=f"<b>{row['municipio']}</b><br/>Zona: {row['ZONA']}<br/>Entregas: {qtd}",
                max_width=260
            ),
            tooltip=f"{row['municipio']} • {row['ZONA']} • {qtd} entrega(s)",
        ).add_to(m)

    if lats and lons:
        m.fit_bounds([[min(lats), min(lons)], [max(lats), max(lons)]])

    legend_html = """
    <div style="
        position: fixed; bottom: 20px; left: 20px; z-index: 9999;
        background: rgba(255,255,255,0.98); padding: 10px 12px;
        border: 1px solid #bbb; border-radius: 8px; color:#111; font-size: 13px;">
      <div style="font-weight:700; margin-bottom:6px;">Legenda</div>
      <div><span style="background:red; width:12px; height:12px; display:inline-block; border-radius:50%; margin-right:6px;"></span>Capital</div>
      <div><span style="background:blue; width:12px; height:12px; display:inline-block; border-radius:50%; margin-right:6px;"></span>Metropolitana</div>
      <div><span style="background:gray; width:12px; height:12px; display:inline-block; border-radius:50%; margin-right:6px;"></span>Interior</div>
      
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))
    st_folium(m, width=None, height=PLOT_HEIGHT)


# =========================================================
# =====================  NORSA (XML)  =====================
# =========================================================
def _detect_ns(root: ET.Element) -> Dict[str, str]:
    m = re.match(r"\{(.+)\}", root.tag or "")
    return {"nfe": m.group(1)} if m else {"nfe": ""}


def _find_text(parent: ET.Element, xpath: str, ns: Dict[str, str]) -> Optional[str]:
    if parent is None:
        return None
    if ns.get("nfe"):
        el = parent.find(xpath, ns)
    else:
        el = parent.find(xpath.replace("nfe:", ""))
    if el is not None and el.text:
        t = el.text.strip()
        return t if t else None
    return None


def _find_first_text(parent: ET.Element, xpaths: List[str], ns: Dict[str, str]) -> Optional[str]:
    for xp in xpaths:
        v = _find_text(parent, xp, ns)
        if v is not None:
            return v
    return None


def _to_float(s: Optional[str]) -> Optional[float]:
    if s is None:
        return None
    try:
        return float(s)
    except Exception:
        return None


def parse_nfe_xml(xml_bytes: bytes) -> Dict[str, Any]:
    root = ET.fromstring(xml_bytes)
    ns = _detect_ns(root)

    inf = root.find(".//nfe:infNFe", ns) if ns.get("nfe") else root.find(".//infNFe")
    if inf is None:
        raise ValueError("Não encontrei <infNFe> no XML. Verifique se é NF-e válida.")

    numero = _find_text(inf, "./nfe:ide/nfe:nNF", ns)
    nome = _find_text(inf, "./nfe:dest/nfe:xNome", ns)

    xLgr = _find_text(inf, "./nfe:dest/nfe:enderDest/nfe:xLgr", ns)
    nro  = _find_text(inf, "./nfe:dest/nfe:enderDest/nfe:nro", ns)
    xCpl = _find_text(inf, "./nfe:dest/nfe:enderDest/nfe:xCpl", ns)

    endereco_parts = [p for p in [xLgr, nro] if p]
    endereco = ", ".join(endereco_parts) if endereco_parts else None
    if xCpl:
        endereco = f"{endereco} - {xCpl}" if endereco else xCpl

    bairro = _find_text(inf, "./nfe:dest/nfe:enderDest/nfe:xBairro", ns)
    municipio = _find_text(inf, "./nfe:dest/nfe:enderDest/nfe:xMun", ns)
    cep = _find_text(inf, "./nfe:dest/nfe:enderDest/nfe:CEP", ns)

    fone = _find_first_text(
        inf,
        [
            "./nfe:dest/nfe:enderDest/nfe:fone",
            "./nfe:emit/nfe:enderEmit/nfe:fone",
        ],
        ns,
    )

    v_nf = _to_float(_find_text(inf, "./nfe:total/nfe:ICMSTot/nfe:vNF", ns))
    q_vol = _to_float(_find_text(inf, "./nfe:transp/nfe:vol/nfe:qVol", ns))

    dets = inf.findall("./nfe:det", ns) if ns.get("nfe") else inf.findall("./det")
    sum_qcom = 0.0
    for det in dets:
        qcom = _to_float(_find_text(det, "./nfe:prod/nfe:qCom", ns))
        if qcom is not None:
            sum_qcom += qcom

    quantidade = q_vol if q_vol is not None else (sum_qcom if sum_qcom > 0 else None)

    peso_bruto = _to_float(
        _find_first_text(
            inf,
            [
                "./nfe:transp/nfe:vol/nfe:pesoB",
                "./nfe:transp/nfe:vol/nfe:pBruto",
            ],
            ns,
        )
    )

    return {
        "Nº": numero,
        "NOME / RAZÃO SOCIAL": nome,
        "ENDEREÇO": endereco,
        "BAIRRO / DISTRITO": bairro,
        "MUNICÍPIO": municipio,
        "CEP": cep,
        "V. TOTAL DA NOTA": v_nf,
        "QUANTIDADE": quantidade,
        "PESO BRUTO": peso_bruto,
        "TELEFONE / FAX": fone,
    }


def salvar_excel_bytes_norsa(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="NFes")
    output.seek(0)
    return output.read()


# =========================================================
# =====================  UI PRINCIPAL  ====================
# =========================================================
st.write("")
cliente = st.selectbox("👤 Selecione o cliente", CLIENTES, index=0)

st.write("")
st.divider()

if cliente.startswith("PLUMA"):
    st.subheader("📌 PLUMA ESPUMAS LTDA — Upload TXT")
    st.caption("Envie o `texto_extraido.txt`. Eu extraio os campos, classifico zona (IBGE) e calculo frete + mapa.")

    uploaded = st.file_uploader("Selecione o arquivo TXT", type=["txt"], accept_multiple_files=False)

    if uploaded is None:
        st.info("Faça o upload do TXT para iniciar a extração.")
        st.stop()

    texto = uploaded.read().decode("utf-8", errors="ignore")

    with st.spinner("Extraindo dados do TXT..."):
        registros = extrair_notas_de_texto(texto)

    if not registros:
        st.warning("Não encontrei notas no TXT. Verifique o layout.")
        st.stop()

    origem = detectar_origem_por_municipios(registros)
    st.caption(f"🏁 Origem detectada automaticamente: **{origem}**")

    excel_bytes, df = salvar_excel_bytes_pluma(registros, origem)

    capital_count = int((df["ZONA"] == "Capital").sum())
    metro_count = int((df["ZONA"] == "Metropolitana").sum())
    interior_count = int((~df["ZONA"].isin(["Capital", "Metropolitana"])).sum())

    col1, col2, col3 = st.columns(3)
    col1.markdown(
        "<div class='kpi'><div class='label'>Notas extraídas</div>"
        f"<div class='value'>{len(df)}</div></div>",
        unsafe_allow_html=True
    )
    col2.markdown(
        "<div class='kpi'><div class='label'>Municípios distintos</div>"
        f"<div class='value'>{df['MUNICÍPIO'].nunique()}</div></div>",
        unsafe_allow_html=True
    )
    col3.markdown(
        "<div class='kpi'><div class='label'>Capital / Metro / Interior</div>"
        f"<div class='value'>{capital_count}/{metro_count}/{interior_count}</div></div>",
        unsafe_allow_html=True
    )

    st.write("")
    st.download_button(
        label="⬇️ Baixar Excel extraído (PLUMA)",
        data=excel_bytes,
        file_name="pluma_notas_extraidas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Pré-visualização (primeiras linhas)"):
        st.dataframe(df.head(30), use_container_width=True)

    st.markdown("#### 🗺️ Mapa ilustrativo dos destinos (PLUMA)")
    cache_col1, _ = st.columns([1, 3])
    with cache_col1:
        if st.button("🧹 Limpar cache de coordenadas"):
            st.session_state["geocache"] = {}
            try:
                if os.path.exists(GEOCACHE_FILE):
                    os.remove(GEOCACHE_FILE)
            except Exception:
                pass
            st.success("Cache limpo! Gere o mapa novamente.")

    if "geocache" not in st.session_state:
        st.session_state["geocache"] = load_geocache()

    municipios_series = (
        df["MUNICÍPIO"]
        .dropna()
        .astype(str)
        .map(sanitize_municipio_name)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
    )

    contagem_municipios = municipios_series.value_counts().to_dict()
    municipios = list(contagem_municipios.keys())

    origem_norm = strip_accents_upper(origem)
    uf_origem = "AL" if "MACEIO" in origem_norm else "PE"

    pontos, nao_plotados = [], []
    for mun in municipios:
        z, _ = classificar_zona_ibge(mun, origem)
        latlon = geocode_city(mun, uf_origem)
        if latlon:
            lat, lon = latlon
            pontos.append({
                "municipio": f"{mun}, {uf_origem}",
                "lat": lat,
                "lon": lon,
                "ZONA": z,
                "QTD_ENTREGAS": int(contagem_municipios.get(mun, 1))
            })
        else:
            nao_plotados.append(mun)

    df_map = pd.DataFrame(pontos)
    st.caption(f"🗺️ Plotados: {len(df_map)} | Municípios distintos no TXT: {len(municipios)}")
    build_map_folium(df_map, origem)

    if nao_plotados:
        st.caption("⚠️ Municípios não plotados (sem coordenadas): " + ", ".join(sorted(set(nao_plotados))))

else:
    st.subheader("📌 NORSA REFRIGERANTES S.A — Upload XML")
    st.caption("Envie um ou vários XMLs de NF-e. Eu extraio os campos e gero uma planilha única.")

    files = st.file_uploader("Arraste aqui seus XMLs (pode mandar vários)", type=["xml"], accept_multiple_files=True)

    if not files:
        st.info("Envie os XMLs acima para começar.")
        st.stop()

    rows, erros = [], []
    for f in files:
        try:
            row = parse_nfe_xml(f.getvalue())
            row["_arquivo"] = f.name
            rows.append(row)
        except Exception as e:
            erros.append({"arquivo": f.name, "erro": str(e)})

    if erros:
        st.warning("Alguns arquivos não puderam ser processados.")
        st.dataframe(pd.DataFrame(erros), use_container_width=True)

    if rows:
        df = pd.DataFrame(rows)
        cols = [
            "Nº",
            "NOME / RAZÃO SOCIAL",
            "ENDEREÇO",
            "BAIRRO / DISTRITO",
            "MUNICÍPIO",
            "CEP",
            "V. TOTAL DA NOTA",
            "QUANTIDADE",
            "PESO BRUTO",
            "TELEFONE / FAX",
            "_arquivo",
        ]
        df = df[[c for c in cols if c in df.columns]]

        st.subheader("✅ Resultado (NORSA)")
        st.dataframe(df, use_container_width=True)

        excel_bytes = salvar_excel_bytes_norsa(df)
        st.download_button(
            "⬇️ Baixar Excel (.xlsx) (NORSA)",
            data=excel_bytes,
            file_name="norsa_extracao_nfe.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

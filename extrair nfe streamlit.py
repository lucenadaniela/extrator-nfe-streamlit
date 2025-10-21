

import io
import os
import re
import json
import time
import unicodedata
import pandas as pd
import streamlit as st

# --- geocoding opcional ---
try:
    from geopy.geocoders import Nominatim
    _GEOPY_OK = True
except Exception:
    _GEOPY_OK = False

import folium
from streamlit_folium import st_folium

# =================== CONFIG ===================
APP_TITLE = "📄 Extrator de NF-e → Excel"
SUBTITLE  = "Classificação IBGE (Capital/Metropolitana) + Mapa dos Destinos em PE"
GEOCACHE_FILE = "geocache_pe.json"
PLOT_HEIGHT = 520

# Base IBGE (RMR/PE) — normalizada
CAPITAL_IBGE = {"RECIFE"}
RMR_IBGE = {
    "RECIFE", "OLINDA", "JABOATAO DOS GUARARAPES", "PAULISTA",
    "CABO DE SANTO AGOSTINHO", "IPOJUCA", "CAMARAGIBE",
    "SAO LOURENCO DA MATA", "ABREU E LIMA", "IGARASSU",
    "ITAPISSUMA", "ARACOIABA", "MORENO", "ILHA DE ITAMARACA",
}

VALOR_FRETE_CAPITAL = 165.00
VALOR_FRETE_METRO   = 170.50

# Regex auxiliares
NUM_BR = r'(\d{1,3}(?:\.\d{3})*(?:,\d+)?|\d+(?:,\d+)?)'
TEL_PAT = re.compile(r'(\(?\d{2}\)?\s*\d{4,5}[-\s]?\d{4}|\b\d{10,11}\b)')
RE_CEP_EXATO = re.compile(r'\b\d{5}-\d{3}\b')

# =================== ESTILO ===================
st.set_page_config(page_title="Extrator NF-e", page_icon="📄", layout="wide")
st.markdown("""
<style>
.main .block-container {max-width: 1200px;}
.stDownloadButton > button {
    border-radius: 10px; padding: 10px 16px; font-weight: 600;
}
.kpi {text-align:center; border-radius:14px; padding:16px; border:1px solid #ececf3; background:#fff;}
.kpi .label {color:#4b5563; font-size:12px;}
.kpi .value {font-weight:800; font-size:24px; color:#111827;}
.leaflet-control-attribution {font-size:12px !important;}
</style>
""", unsafe_allow_html=True)

# =================== HELPERS ===================
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
    s = re.sub(r"\b(PE|PERNAMBUCO)\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"\s{2,}", " ", s).strip()
    return s

def classificar_zona_ibge(municipio: str):
    key = strip_accents_upper(municipio)
    if key in CAPITAL_IBGE:
        return "Capital", VALOR_FRETE_CAPITAL
    if key in RMR_IBGE:
        return "Metropolitana", VALOR_FRETE_METRO
    return "Outros", None

def prox_nao_vazia(linhas, j, max_look=15):
    n = len(linhas); k = j; passos = 0
    while k < n and passos < max_look:
        val = (linhas[k] or "").strip()
        if val: return val
        k += 1; passos += 1
    return ""

def normalize_phone(raw: str) -> str:
    if not raw: return ""
    digits = re.sub(r"\D", "", raw)
    if len(digits) == 11: return f"({digits[:2]}){digits[2:7]}-{digits[7:]}"
    if len(digits) == 10: return f"({digits[:2]}){digits[2:6]}-{digits[6:]}"
    return re.sub(r"\s+", " ", raw).strip()

def find_phone_in_text(txt: str) -> str:
    m = TEL_PAT.search(txt or "")
    return normalize_phone(m.group(1)) if m else ""

def extrair_cep_exato(txt: str) -> str:
    if not txt: return ""
    m = RE_CEP_EXATO.search(txt)
    return m.group(0) if m else ""

# =================== EXTRAÇÃO ===================
def extrair_notas_de_texto(texto: str):
    linhas = [l.strip() for l in texto.splitlines()]
    registros, atual = [], None
    n, i = len(linhas), 0
    while i < n:
        linha = linhas[i]

        # Início da NF
        m_num = re.search(r'\bN[ºo]\s*([0-9\.\-]+)', linha, flags=re.IGNORECASE)
        if m_num:
            if atual: registros.append(atual)
            atual = {
                "Nº": m_num.group(1),
                "NOME / RAZÃO SOCIAL": None,
                "ENDEREÇO": None,
                "BAIRRO / DISTRITO": None,
                "MUNICÍPIO": None,
                "CEP": None,
                "VALOR TOTAL DA NOTA": None,
                "QUANTIDADE": None,          # << NOVO
                "PESO BRUTO": None,
                "TELEFONE / FAX": None,
                "TELEFONE 2": None,
            }
            i += 1; continue

        if atual is not None:
            # NOME / RAZÃO SOCIAL
            if re.fullmatch(r'NOME\s*/\s*RAZÃO SOCIAL', linha, flags=re.IGNORECASE) and atual["NOME / RAZÃO SOCIAL"] is None:
                v = prox_nao_vazia(linhas, i + 1)
                if v: atual["NOME / RAZÃO SOCIAL"] = v

            # ENDEREÇO
            if re.fullmatch(r'ENDEREÇO', linha, flags=re.IGNORECASE):
                v = prox_nao_vazia(linhas, i + 1)
                if v: atual["ENDEREÇO"] = v

            # BAIRRO / DISTRITO
            if re.fullmatch(r'BAIRRO\s*/\s*DISTRITO', linha, flags=re.IGNORECASE):
                v = prox_nao_vazia(linhas, i + 1)
                if v: atual["BAIRRO / DISTRITO"] = v

            # MUNICÍPIO
            if re.fullmatch(r'MUNICÍPIO', linha, flags=re.IGNORECASE):
                v = prox_nao_vazia(linhas, i + 1)
                if v:
                    v_limpo = re.split(r'\bUF\b|CEP|\bPE\b|\d{2}:\d{2}:\d{2}', v, maxsplit=1)[0].strip(" -")
                    atual["MUNICÍPIO"] = v_limpo or v

            # VALOR TOTAL DA NOTA
            if re.fullmatch(r'VALOR TOTAL DA NOTA', linha, flags=re.IGNORECASE):
                v = prox_nao_vazia(linhas, i + 1)
                if v:
                    m_val = re.search(NUM_BR, v)
                    atual["VALOR TOTAL DA NOTA"] = (m_val.group(1) if m_val else v).replace("R$", "").strip()

            # ===== PESO BRUTO =====
            if re.fullmatch(r'PESO\s+BRUTO', linha, flags=re.IGNORECASE) and not atual["PESO BRUTO"]:
                v = prox_nao_vazia(linhas, i + 1, max_look=10)
                if v:
                    m = re.search(NUM_BR, v)
                    if m: atual["PESO BRUTO"] = m.group(1)
            if not atual["PESO BRUTO"]:
                m_inline = re.search(r'\bPESO\s+BRUTO\b.*?' + NUM_BR, linha, flags=re.IGNORECASE)
                if m_inline: atual["PESO BRUTO"] = m_inline.group(1)
            if not atual["PESO BRUTO"] and ("VOLUME" in linha.upper() or "NUMERAÇÃO" in linha.upper()):
                m = re.search(NUM_BR + r'(?!.*\d)', linha)
                if m: atual["PESO BRUTO"] = m.group(1)

            # ===== QUANTIDADE =====
            # Captura "2 VOLUMES" na mesma linha OU abaixo do rótulo "QUANTIDADE"
            if atual["QUANTIDADE"] is None:
                # 1) mesma linha: "... 2 VOLUMES ..."
                m_qt = re.search(r'(\d+)\s+VOLUMES', linha, flags=re.IGNORECASE)
                if m_qt:
                    atual["QUANTIDADE"] = m_qt.group(1)
                else:
                    # 2) bloco com rótulo na linha e número na próxima
                    if re.search(r'\bQUANTIDADE\b', linha, flags=re.IGNORECASE):
                        v = prox_nao_vazia(linhas, i + 1, max_look=5)
                        # tenta "2 VOLUMES" ou apenas um número sozinho
                        m_next = re.search(r'(\d+)\s+VOLUMES', v, flags=re.IGNORECASE) or re.search(r'\b(\d+)\b', v)
                        if m_next:
                            atual["QUANTIDADE"] = m_next.group(1)

            # ===== CEP =====
            if linha.strip().upper() == "CEP" and not atual["CEP"]:
                v = prox_nao_vazia(linhas, i + 1, max_look=15); atual["CEP"] = extrair_cep_exato(v)
            elif not atual["CEP"]:
                atual["CEP"] = extrair_cep_exato(linha)

            # ===== TELEFONE / FAX =====
            if re.fullmatch(r'TELEFONE\s*/\s*FAX', linha, flags=re.IGNORECASE) and not atual["TELEFONE / FAX"]:
                v = prox_nao_vazia(linhas, i + 1, max_look=6)
                if not re.match(r'INSCRIÇÃO|DESTINAT[ÁA]RIO', v, flags=re.IGNORECASE):
                    tel = find_phone_in_text(v) or find_phone_in_text(linha)
                    if tel: atual["TELEFONE / FAX"] = tel
            if not atual["TELEFONE / FAX"] and "TELEFONE / FAX" in linha.upper():
                if not re.search(r'INSCRIÇÃO\s+ESTADUAL', linha, flags=re.IGNORECASE):
                    tel = find_phone_in_text(linha)
                    if tel: atual["TELEFONE / FAX"] = tel

            # ===== TELEFONE 2 =====
            if ("RASTREAMENTO" in linha.upper()) or ("ENTREGAID" in linha.upper()) or ("TELEFONE 2" in linha.upper()):
                m_t2 = re.search(r'TELEFONE\s*2\s*:\s*([^;]*)', linha, flags=re.IGNORECASE)
                if m_t2:
                    valor = m_t2.group(1).strip()
                    if valor: atual["TELEFONE 2"] = find_phone_in_text(valor)

        i += 1

    if atual: registros.append(atual)
    return registros

def salvar_excel_bytes(registros) -> tuple[bytes, pd.DataFrame]:
    df = pd.DataFrame(registros)

    # Une linhas quebradas com mesmo Nº consecutivo
    df_clean, skip_next = [], False
    for i in range(len(df)):
        if skip_next:
            skip_next = False
            continue
        row = df.iloc[i]
        if i + 1 < len(df) and str(df.iloc[i + 1]["Nº"]) == str(row["Nº"]):
            combined = df.iloc[i + 1].copy()
            combined["Nº"] = row["Nº"]
            for campo in [
                "PESO BRUTO", "QUANTIDADE", "TELEFONE / FAX", "TELEFONE 2",
                "NOME / RAZÃO SOCIAL", "ENDEREÇO",
                "BAIRRO / DISTRITO", "MUNICÍPIO", "CEP",
                "VALOR TOTAL DA NOTA",
            ]:
                if not str(combined.get(campo, "")).strip():
                    combined[campo] = row.get(campo)
            df_clean.append(combined); skip_next = True
        else:
            df_clean.append(row)
    df = pd.DataFrame(df_clean).reset_index(drop=True)

    # Classificação IBGE
    zonas, fretes = [], []
    for mun in df["MUNICÍPIO"].fillna(""):
        z, f = classificar_zona_ibge(mun)
        zonas.append(z); fretes.append(f)
    df["ZONA"] = zonas; df["VALOR FRETE"] = fretes

    colunas = [
        "Nº", "NOME / RAZÃO SOCIAL", "ENDEREÇO",
        "BAIRRO / DISTRITO", "MUNICÍPIO", "CEP",
        "VALOR TOTAL DA NOTA", "QUANTIDADE", "PESO BRUTO",
        "TELEFONE / FAX", "TELEFONE 2",
        "ZONA", "VALOR FRETE",
    ]
    df = df.reindex(columns=colunas)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Notas")
    buffer.seek(0)
    return buffer.read(), df

# =================== GEOCACHE ===================
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

# ===== Fallback de coordenadas (RMR + variações) =====
COORDS_FALLBACK_RAW = {
    "RECIFE, PE": (-8.0476, -34.8770),
    "OLINDA, PE": (-8.0101, -34.8545),
    "JABOATÃO DOS GUARARAPES, PE": (-8.1120, -35.0140),
    "PAULISTA, PE": (-7.9400, -34.8731),
    "CABO DE SANTO AGOSTINHO, PE": (-8.2822, -35.0320),
    "IPOJUCA, PE": (-8.3983, -35.0639),
    "CAMARAGIBE, PE": (-8.0207, -34.9786),
    "SÃO LOURENÇO DA MATA, PE": (-8.0062, -35.0199),
    "ABREU E LIMA, PE": (-7.9007, -34.9027),
    "IGARASSU, PE": (-7.8340, -34.9069),
    "ITAPISSUMA, PE": (-7.7750, -34.8954),
    "ARAÇOIABA, PE": (-7.7913, -35.0800),
    "MORENO, PE": (-8.1180, -35.0920),
    "ILHA DE ITAMARACÁ, PE": (-7.7478, -34.8332),
    # variações
    "JABOATAO DOS GUARARAPES, PE": (-8.1120, -35.0140),
    "ILHA DE ITAMARACA, PE": (-7.7478, -34.8332),
    "SAO LOURENCO DA MATA, PE": (-8.0062, -35.0199),
    "PAULISTA, PERNAMBUCO": (-7.9400, -34.8731),
}
def _norm_place_key(s: str) -> str:
    return strip_accents_upper(s)
COORDS_FALLBACK_NORM = { _norm_place_key(k): v for k, v in COORDS_FALLBACK_RAW.items() }

def geocode_city(city: str) -> tuple | None:
    """
    Retorna (lat, lon) para 'city, PE'.
    Usa cache; tenta fallback normalizado; se possível, Nominatim.
    Ignora cache negativo (None) e re-tenta.
    """
    city_clean = sanitize_municipio_name(city)
    key_raw  = f"{city_clean}, PE"
    key_norm = _norm_place_key(key_raw)

    cache = st.session_state.get("geocache", {})

    # coordenadas válidas em cache
    if key_raw in cache and cache[key_raw]:
        return tuple(cache[key_raw])

    # fallback normalizado
    if key_norm in COORDS_FALLBACK_NORM:
        latlon = COORDS_FALLBACK_NORM[key_norm]
        cache[key_raw] = latlon; st.session_state["geocache"] = cache; save_geocache(cache)
        return latlon

    # Nominatim
    if _GEOPY_OK:
        try:
            geolocator = Nominatim(user_agent="nf_extractor_ws")
            loc = geolocator.geocode(f"{city_clean}, Pernambuco, Brazil", timeout=10)
            if loc:
                latlon = (loc.latitude, loc.longitude)
                cache[key_raw] = latlon; st.session_state["geocache"] = cache; save_geocache(cache)
                time.sleep(1.0)
                return latlon
        except Exception:
            pass

    # marca falha agora
    cache[key_raw] = None; st.session_state["geocache"] = cache; save_geocache(cache)
    return None

# =================== MAPA (FOLIUM) ===================
def _color_for(z):
    if z == "Capital":
        return "red"
    if z == "Metropolitana":
        return "blue"
    return "gray"


def build_map_folium(df_destinos: pd.DataFrame):
    """Plota destinos (municípios de PE) com Folium; tamanho do ponto ~ nº de entregas."""
    if df_destinos.empty:
        st.info("Sem destinos válidos para plotar no mapa.")
        return

    m = folium.Map(location=[-8.38, -37.86], zoom_start=6, tiles="OpenStreetMap")

    lats, lons = [], []

    def radius_for(n):
        base = 6
        extra = min(n, 14)
        return base + extra

    for _, row in df_destinos.iterrows():
        lat, lon = row["lat"], row["lon"]
        lats.append(lat)
        lons.append(lon)
        entregas = int(row.get("entregas", 1))

        folium.CircleMarker(
            location=[lat, lon],
            radius=radius_for(entregas),
            weight=2,
            color=_color_for(row["ZONA"]),
            fill=True,
            fill_color=_color_for(row["ZONA"]),
            fill_opacity=0.9,
            popup=folium.Popup(
                html=(
                    f"<b>{row['municipio']}</b>"
                    f"<br/>Zona: {row['ZONA']}"
                    f"<br/>Entregas: {entregas}"
                ),
                max_width=260,
            ),
            tooltip=f"{row['municipio']} • {row['ZONA']} • Entregas: {entregas}",
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
      <div><span style="background:gray; width:12px; height:12px; display:inline-block; border-radius:50%; margin-right:6px;"></span>Outros</div>
      <div style="margin-top:6px; font-size:12px; color:#333;">
        • O tamanho do ponto indica o nº de entregas
      </div>
    </div>
    """
    m.get_root().html.add_child(folium.Element(legend_html))
    st_folium(m, width=None, height=PLOT_HEIGHT)



# =================== UI ===================
st.markdown(f"### {APP_TITLE}")
st.caption(SUBTITLE)
st.write("")

left, right = st.columns([1.2, 1])
with left:
    st.markdown("#### 1) Envie o arquivo `.txt`")
    uploaded = st.file_uploader("Selecione o `texto_extraido.txt`", type=["txt"])
with right:
    st.empty()

st.write("")

if uploaded is not None:
    texto = uploaded.read().decode("utf-8", errors="ignore")
    with st.spinner("Extraindo dados do TXT..."):
        registros = extrair_notas_de_texto(texto)

    if not registros:
        st.warning("Não encontrei notas no TXT. Verifique o layout.")
        st.stop()

    excel_bytes, df = salvar_excel_bytes(registros)

    # KPIs
    total_frete = float(pd.to_numeric(df["VALOR FRETE"], errors="coerce").fillna(0).sum())
    capital = int((df["ZONA"] == "Capital").sum())
    metro   = int((df["ZONA"] == "Metropolitana").sum())
    outros  = int((df["ZONA"] == "Outros").sum())
    
    col1, col2, col3, col4 = st.columns(4)
    col1.markdown("<div class='kpi'><div class='label'>Notas extraídas</div>"
                  f"<div class='value'>{len(df)}</div></div>", unsafe_allow_html=True)
    col2.markdown("<div class='kpi'><div class='label'>Municípios distintos</div>"
                  f"<div class='value'>{df['MUNICÍPIO'].nunique()}</div></div>", unsafe_allow_html=True)
    col3.markdown("<div class='kpi'><div class='label'>Capital/Metropolitana/Outros</div>"
                  f"<div class='value'>{capital}/{metro}/{outros}</div></div>", unsafe_allow_html=True)
    col4.markdown("<div class='kpi'><div class='label'>Valor total de frete (R$)</div>"
                  f"<div class='value'>{total_frete:,.2f}</div></div>", unsafe_allow_html=True)

    st.write("")
    st.download_button(
        label="⬇️ Baixar Excel extraído",
        data=excel_bytes,
        file_name="notas_extraidas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("Pré-visualização (primeiras linhas)"):
        st.dataframe(df.head(20), use_container_width=True)

    # ----- MAPA -----
    st.markdown("#### 2) Mapa ilustrativo dos destinos em Pernambuco")

    # Limpar cache (caso coordenadas antigas erradas)
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

    # Contagem de entregas por município (sanitizado)
    mun_series = (
        df["MUNICÍPIO"]
        .dropna()
        .astype(str)
        .map(sanitize_municipio_name)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
    )
    
    entregas_por_mun = mun_series.value_counts()          # Series: mun -> contagem
    municipios = entregas_por_mun.index.tolist()
    
    pontos = []
    nao_plotados = []
    for mun in municipios:
        z, _ = classificar_zona_ibge(mun)
        latlon = geocode_city(mun)
        if latlon:
            lat, lon = latlon
            pontos.append({
                "municipio": f"{mun}, PE",
                "lat": lat,
                "lon": lon,
                "ZONA": z,
                "entregas": int(entregas_por_mun[mun]),
            })
        else:
            nao_plotados.append(mun)


        df_map = pd.DataFrame(pontos)
    st.caption(f"🗺️ Plotados: {len(df_map)} | Municípios distintos no TXT: {len(entregas_por_mun)}")
    build_map_folium(df_map)

    if nao_plotados:
        st.caption("⚠️ Municípios não plotados (sem coordenadas/OSM): " + ", ".join(sorted(set(nao_plotados))))

else:
    st.info("Faça o upload do arquivo TXT para iniciar a extração.")

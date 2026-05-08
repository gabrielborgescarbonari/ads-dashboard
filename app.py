import sys
import os
sys.stdout.reconfigure(encoding="utf-8")

# Injeta secrets do Streamlit Cloud em os.environ (funciona local e na nuvem)
try:
    import streamlit as st
    for _k, _v in st.secrets.items():
        os.environ.setdefault(_k, str(_v))
except Exception:
    pass

from datetime import date, timedelta

import pandas as pd
import streamlit as st

from src import google_ads_client, meta_ads

st.set_page_config(page_title="Dashboard de Ads", layout="wide")
st.title("Dashboard de Ads — Google + Meta")

# --- Sidebar ---
st.sidebar.header("Filtros")

col1, col2 = st.sidebar.columns(2)
with col1:
    date_start = st.date_input("Inicio", value=date.today() - timedelta(days=30))
with col2:
    date_end = st.date_input("Fim", value=date.today())

plataformas = st.sidebar.multiselect(
    "Plataforma",
    options=["Meta", "Google"],
    default=["Meta", "Google"],
)

carregar = st.sidebar.button("Carregar dados", type="primary", use_container_width=True)


@st.cache_data(ttl=3600, show_spinner="Buscando dados nas APIs...")
def load_data(start: str, end: str) -> pd.DataFrame:
    frames = []

    try:
        meta_df = meta_ads.fetch_insights(start, end)
        if not meta_df.empty:
            frames.append(meta_df)
    except Exception as e:
        st.warning(f"Meta Ads: {e}")

    try:
        google_df = google_ads_client.fetch_insights(start, end)
        if not google_df.empty:
            frames.append(google_df)
    except Exception as e:
        st.warning(f"Google Ads: {e}")

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


if carregar:
    st.session_state["loaded"] = True
    st.cache_data.clear()

if "loaded" not in st.session_state:
    st.info("Configure os filtros e clique em 'Carregar dados'.")
    st.stop()

df = load_data(str(date_start), str(date_end))

if df.empty:
    st.warning("Nenhum dado encontrado para o periodo selecionado.")
    st.stop()

# --- Filtros dinamicos ---
if plataformas:
    df = df[df["plataforma"].isin(plataformas)]

campanhas = st.sidebar.multiselect("Campanha", options=sorted(df["campanha"].dropna().unique()))
if campanhas:
    df = df[df["campanha"].isin(campanhas)]

conjuntos = st.sidebar.multiselect("Conjunto / Ad Group", options=sorted(df["conjunto"].dropna().unique()))
if conjuntos:
    df = df[df["conjunto"].isin(conjuntos)]

anuncios = st.sidebar.multiselect("Anuncio", options=sorted(df["anuncio"].dropna().unique()))
if anuncios:
    df = df[df["anuncio"].isin(anuncios)]

st.sidebar.markdown("**UTMs**")
for utm_field, label in [
    ("utm_source", "utm_source"),
    ("utm_medium", "utm_medium"),
    ("utm_campaign", "utm_campaign"),
    ("utm_content", "utm_content"),
    ("utm_term", "utm_term"),
]:
    options = sorted(df[utm_field].replace("", pd.NA).dropna().unique())
    selected = st.sidebar.multiselect(label, options=options)
    if selected:
        df = df[df[utm_field].isin(selected)]

# --- Metricas resumo ---
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Investimento", f"R$ {df['investimento'].sum():,.2f}")
c2.metric("Impressoes", f"{int(df['impressoes'].sum()):,}")
c3.metric("Cliques", f"{int(df['cliques'].sum()):,}")
c4.metric("Conversoes", f"{df['conversoes'].sum():,.0f}")
c5.metric("CPA medio", f"R$ {(df['investimento'].sum() / df['conversoes'].sum()):,.2f}" if df['conversoes'].sum() > 0 else "R$ -")

st.markdown("---")

# --- Tabela ---
display_df = df[[
    "plataforma", "data", "campanha", "conjunto", "anuncio",
    "investimento", "impressoes", "cliques", "ctr", "cpc", "cpm",
    "conversoes", "cpa",
    "utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term",
]].copy()

display_df["data"] = pd.to_datetime(display_df["data"]).dt.strftime("%d/%m/%Y")

display_df.columns = [
    "Plataforma", "Data", "Campanha", "Conjunto", "Anuncio",
    "Investimento (R$)", "Impressoes", "Cliques", "CTR (%)", "CPC (R$)", "CPM (R$)",
    "Conversoes", "CPA (R$)",
    "UTM Source", "UTM Medium", "UTM Campaign", "UTM Content", "UTM Term",
]

st.dataframe(
    display_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Investimento (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
        "CPC (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
        "CPM (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
        "CPA (R$)": st.column_config.NumberColumn(format="R$ %.2f"),
        "CTR (%)": st.column_config.NumberColumn(format="%.2f%%"),
        "Conversoes": st.column_config.NumberColumn(format="%.0f"),
    },
)

st.caption(f"{len(display_df):,} linhas exibidas")

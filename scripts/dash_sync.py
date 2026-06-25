# -*- coding: utf-8 -*-
"""Orquestrador do sync que alimenta o dash de midia a partir das APIs.

Constroi um cache local (data/dash_data.xlsx) com abas Google e Meta no formato
de 17 colunas que o gerar_dashboard_midia.py ja le. Publico e Formato sao
derivados do nome da campanha (dash_derive). Conversoes: Google = metrics.conversions;
Meta = lead + conversa de WhatsApp iniciada (definicao calibrada contra a planilha).

Modos:
  python dash_sync.py --full            puxa todo 2026 das APIs e reconstroi o cache
  python dash_sync.py --from-raw        reconstroi o cache a partir de data/raw_*.csv (sem API)
  python dash_sync.py --incremental [N] atualiza os ultimos N dias (default 7) no cache
"""
import os
import sys
import json
import argparse
from datetime import date, timedelta

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dash_fetch as fetch
from dash_derive import derive_publico, derive_formato, precisa_revisao

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
CACHE_XLSX = os.path.join(DATA_DIR, "dash_data.xlsx")
RAW_GOOGLE = os.path.join(DATA_DIR, "raw_google.csv")
RAW_META = os.path.join(DATA_DIR, "raw_meta.csv")
START_HIST = "2026-01-01"

# Historico de Meta da planilha (GROWTH) usado ate o corte, porque traz a conversao
# de SITE do passado que a API interina nao tem (conversoes personalizadas nao sao
# retroativas). Ate CUTOVER_META: Meta vem do Excel; depois: vem da API.
HIST_META_XLSX = os.path.join(DATA_DIR, "GROWTH_Estudo plataformas de mídia (2).xlsx")
CUTOVER_META = "2026-06-18"
# Conta GNDI: na planilha nova a aba e nomeada pelo id da conta (ex.: "407816857488424 -> GNDI").
GNDI_ACCT = "407816857488424"

# Conversao Meta = soma destes action_type.
# lead = leadform (on-Meta); messaging = WhatsApp. As duas custom.<ID> sao as
# conversoes personalizadas das UNICAS campanhas ativas hoje (cada uma so dispara no
# seu funil, sem dupla contagem): LQI MINI = PF JORNADA MINI; LQE MINI = PME LP ENCURTADA.
# So tem efeito no periodo da API (Data > CUTOVER_META); o passado vem da planilha.
# Conversao personalizada NAO e retroativa (confirmado 2026-06-17), por isso o corte.
# Ao reativar campanha pausada, criar a custom do evento dela e somar o ID aqui.
META_CONV_ACTIONS = [
    "lead",
    "onsite_conversion.messaging_conversation_started_7d",
    "offsite_conversion.custom.1682278389684149",  # LQI MINI - PF JORNADA MINI
    "offsite_conversion.custom.1010937478182726",  # LQE MINI - PME LP ENCURTADA
    "offsite_conversion.custom.1329453951952181",  # PF JORNADA NANO
]

COLS_OUT = ["Plataforma", "Conta", "Data", "Campanha", "Conjunto", "Publico", "Formato",
            "Investimento (R$)", "Alcance", "Impressoes", "Cliques", "CTR (%)",
            "CPC (R$)", "CPM (R$)", "Leads - plataforma", "CPL - plataforma",
            "Leads - interno", "CPL - interno", "Visualizacoes", "Inscricoes"]
# Layout legado de 17 colunas da planilha GROWTH (sem Conjunto/Visualizacoes/Inscricoes).
COLS_PLANILHA = ["Plataforma", "Conta", "Data", "Campanha", "Publico", "Formato",
                 "Investimento (R$)", "Alcance", "Impressoes", "Cliques", "CTR (%)",
                 "CPC (R$)", "CPM (R$)", "Leads - plataforma", "CPL - plataforma",
                 "Leads - interno", "CPL - interno"]


def meta_conversoes(actions_json: str) -> float:
    d = json.loads(actions_json) if isinstance(actions_json, str) else {}
    return sum(d.get(a, 0.0) for a in META_CONV_ACTIONS)


def _finalize(df: pd.DataFrame, plataforma: str) -> pd.DataFrame:
    """Recebe df cru (pl,ct,dt,ca,iv,al,im,cl,lp) e devolve no formato de 17 colunas."""
    df = df.copy()
    df["Publico"] = df["ca"].apply(derive_publico)
    df["Formato"] = df["ca"].apply(lambda c: derive_formato(c, plataforma))
    # Conta "Hapvida - Social" = YouTube institucional. Classifica o bloco inteiro pelo
    # nome da conta (mais robusto que parsear cada campanha) e roteia a conversao
    # (metrics.conversions = inscricoes) para a metrica Inscricoes, zerando Leads pra
    # nao misturar video com captacao. Visualizacoes vem de metrics.video_views.
    soc = df["ct"].astype(str).str.strip() == "Hapvida - Social"
    df.loc[soc, "Publico"] = "Social"
    df.loc[soc, "Formato"] = "Youtube"
    # Conta Odonto = linha de produto propria; classifica o bloco inteiro pelo nome
    # da conta (as campanhas nao seguem a taxonomia PF/PME). Formato continua vindo
    # do nome (Whatsapp/Leadform/site), entao nao sobrescreve.
    od = df["ct"].astype(str).str.strip() == "Odonto"
    df.loc[od, "Publico"] = "Odonto"
    im = df["im"].astype(float)
    cl = df["cl"].astype(float)
    iv = df["iv"].astype(float)
    conv = df["lp"].astype(float)
    vv = df["vv"].astype(float) if "vv" in df.columns else pd.Series(0.0, index=df.index)
    leads = conv.where(~soc, 0.0)   # captacao (tudo que nao e a conta Social)
    insc = conv.where(soc, 0.0)     # inscricoes (so a conta Social)
    df["CTR (%)"] = (cl / im * 100).where(im > 0, 0).round(2)
    df["CPC (R$)"] = (iv / cl).where(cl > 0, 0).round(2)
    df["CPM (R$)"] = (iv / im * 1000).where(im > 0, 0).round(2)
    cpl = (iv / leads).where(leads > 0, 0).round(2)
    cj = (df["cj"].fillna("(sem conjunto)") if "cj" in df.columns
          else pd.Series("(sem conjunto)", index=df.index))
    out = pd.DataFrame({
        "Plataforma": df["pl"], "Conta": df["ct"], "Data": df["dt"],
        "Campanha": df["ca"], "Conjunto": cj, "Publico": df["Publico"], "Formato": df["Formato"],
        "Investimento (R$)": iv.round(2), "Alcance": df["al"].astype(int),
        "Impressoes": im.astype(int), "Cliques": cl.astype(int),
        "CTR (%)": df["CTR (%)"], "CPC (R$)": df["CPC (R$)"], "CPM (R$)": df["CPM (R$)"],
        "Leads - plataforma": leads.round(2), "CPL - plataforma": cpl,
        "Leads - interno": "", "CPL - interno": "",
        "Visualizacoes": vv.astype(int), "Inscricoes": insc.round(2),
    })
    return out[COLS_OUT]


def _drop_unclassified(df: pd.DataFrame, label: str) -> pd.DataFrame:
    mask = (df["Publico"] == "?") | (df["Formato"] == "?")
    if mask.any():
        n_camp = df.loc[mask, "Campanha"].nunique()
        print(f"  {label}: {int(mask.sum())} linhas / {n_camp} campanha(s) sem "
              f"classificacao removidas do cache (ficam na lista de revisao).")
    return df[~mask].reset_index(drop=True)


def _load_meta_hist() -> pd.DataFrame:
    """Le a aba Meta da planilha GROWTH (ja no formato de 17 colunas) e devolve so
    as linhas ate CUTOVER_META, com Data em ISO. Traz a conversao de site do passado."""
    if not os.path.exists(HIST_META_XLSX):
        print(f"  ! Historico Meta nao encontrado ({HIST_META_XLSX}); usando so a API.")
        return pd.DataFrame(columns=COLS_OUT)
    xl = pd.ExcelFile(HIST_META_XLSX)
    # Layout novo: aba da GNDI nomeada pelo id da conta. Layout legado: aba "Meta".
    sheet = next((s for s in xl.sheet_names if s.strip().startswith(GNDI_ACCT)), None)
    if sheet is None:
        sheet = "Meta" if "Meta" in xl.sheet_names else xl.sheet_names[0]
    h = pd.read_excel(xl, sheet_name=sheet)
    h = h.iloc[:, :17].copy()
    h.columns = COLS_PLANILHA  # planilha nao tem Conjunto; so normaliza acentos
    h["Conta"] = "GNDI"  # a aba Meta da planilha e so GNDI; garante continuidade c/ a API
    h["Conjunto"] = "(sem conjunto)"  # planilha e nivel campanha, sem quebra de conjunto
    h["Visualizacoes"] = 0  # planilha legada nao tem metricas de video
    h["Inscricoes"] = 0
    h["Data"] = pd.to_datetime(h["Data"], errors="coerce")
    h = h[h["Data"] <= pd.Timestamp(CUTOVER_META)].copy()
    h["Data"] = h["Data"].dt.strftime("%Y-%m-%d")
    return h[COLS_OUT]


def _splice_meta(m_api: pd.DataFrame) -> pd.DataFrame:
    """Planilha cobre so a GNDI ate o corte (conversao de site nao e retroativa).
    A API cuida de todo o resto: GNDI depois do corte e as contas de vendas de
    servico em qualquer data (a conversao delas e so WhatsApp, retroativa na API)."""
    hist = _load_meta_hist()  # GNDI, ate o corte
    if hist.empty:
        return m_api
    keep = (m_api["Data"] > CUTOVER_META) | (m_api["Conta"] != "GNDI")
    api_kept = m_api[keep].copy()
    out = pd.concat([hist, api_kept], ignore_index=True)
    print(f"  Meta: planilha GNDI ate {CUTOVER_META} ({len(hist)} linhas) + API "
          f"({len(api_kept)} linhas) = {len(out)} linhas.")
    return out


def build_from_frames(g_raw: pd.DataFrame, m_raw: pd.DataFrame):
    g = _drop_unclassified(_finalize(g_raw, "Google"), "Google")
    m_raw = m_raw.copy()
    m_raw["lp"] = m_raw["actions"].apply(meta_conversoes)
    m = _drop_unclassified(_finalize(m_raw, "Meta"), "Meta")
    m = _splice_meta(m)
    return g, m


def _report_revisao(g_raw, m_raw):
    pend = []
    # A conta Social e classificada por conta (YouTube institucional), nao pelo nome;
    # nao deve aparecer como "sem classificacao".
    soc_camps = set(g_raw.loc[g_raw["ct"].astype(str).str.strip() == "Hapvida - Social", "ca"].unique())
    for ca in g_raw["ca"].unique():
        if ca in soc_camps:
            continue
        if precisa_revisao(ca, "Google"):
            pend.append(("Google", ca))
    for ca in m_raw["ca"].unique():
        if precisa_revisao(ca, "Meta"):
            pend.append(("Meta", ca))
    if pend:
        print(f"\n  ATENCAO: {len(pend)} campanha(s) sem classificacao automatica (revisar):")
        for pl, ca in pend:
            print(f"    [{pl}] {ca}")
    else:
        print("\n  Todas as campanhas classificadas automaticamente (publico/formato).")


def write_cache(g: pd.DataFrame, m: pd.DataFrame):
    os.makedirs(DATA_DIR, exist_ok=True)
    with pd.ExcelWriter(CACHE_XLSX, engine="openpyxl") as xw:
        g.to_excel(xw, sheet_name="Google", index=False)
        m.to_excel(xw, sheet_name="Meta", index=False)
    print(f"\nCache escrito: {CACHE_XLSX}")
    print(f"  Google: {len(g)} linhas | Meta: {len(m)} linhas")
    print(f"  Investimento total: R$ {g['Investimento (R$)'].sum()+m['Investimento (R$)'].sum():,.2f}")


def run_full():
    end = date.today().isoformat()
    print(f"FULL: {START_HIST} -> {end}")
    g_raw = fetch.fetch_google(START_HIST, end)
    g_raw.to_csv(RAW_GOOGLE, index=False, encoding="utf-8")
    m_raw = fetch.fetch_meta(START_HIST, end)
    m_raw.to_csv(RAW_META, index=False, encoding="utf-8")
    g, m = build_from_frames(g_raw, m_raw)
    _report_revisao(g_raw, m_raw)
    write_cache(g, m)


def run_from_raw():
    print("FROM-RAW: reconstruindo cache a partir de data/raw_*.csv")
    g_raw = pd.read_csv(RAW_GOOGLE)
    m_raw = pd.read_csv(RAW_META)
    g, m = build_from_frames(g_raw, m_raw)
    _report_revisao(g_raw, m_raw)
    write_cache(g, m)


def run_incremental(days: int):
    start = (date.today() - timedelta(days=days)).isoformat()
    end = date.today().isoformat()
    print(f"INCREMENTAL: {start} -> {end} (merge no cache existente)")
    g_new = fetch.fetch_google(start, end)
    m_new = fetch.fetch_meta(start, end)
    # atualiza raw substituindo a janela
    g_raw = pd.read_csv(RAW_GOOGLE) if os.path.exists(RAW_GOOGLE) else pd.DataFrame()
    m_raw = pd.read_csv(RAW_META) if os.path.exists(RAW_META) else pd.DataFrame()
    if not g_raw.empty:
        g_raw = g_raw[g_raw["dt"] < start]
    if not m_raw.empty:
        m_raw = m_raw[m_raw["dt"] < start]
    g_raw = pd.concat([g_raw, g_new], ignore_index=True)
    m_raw = pd.concat([m_raw, m_new], ignore_index=True)
    g_raw.to_csv(RAW_GOOGLE, index=False, encoding="utf-8")
    m_raw.to_csv(RAW_META, index=False, encoding="utf-8")
    g, m = build_from_frames(g_raw, m_raw)
    _report_revisao(g_raw, m_raw)
    write_cache(g, m)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--full", action="store_true")
    ap.add_argument("--from-raw", action="store_true")
    ap.add_argument("--incremental", nargs="?", const=7, type=int)
    args = ap.parse_args()
    if args.full:
        run_full()
    elif args.from_raw:
        run_from_raw()
    elif args.incremental is not None:
        run_incremental(args.incremental)
    else:
        ap.print_help()

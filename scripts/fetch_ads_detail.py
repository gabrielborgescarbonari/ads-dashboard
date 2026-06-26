# -*- coding: utf-8 -*-
"""Fetch em nivel de ANUNCIO (Google ad_group_ad + Meta level=ad) para alimentar
o EXPORT de anuncios do dash. NAO mexe no cache principal do dash; gera so
data/ads_detail.csv (pl, ct, dt, ca, cj, ad, ad_id, iv, lp).

O dash continua em nivel de conjunto; este arquivo a parte (servido pelo Vercel)
e carregado sob demanda pelo botao de exportar. Publico/Formato/taxonomia NAO sao
derivados aqui: o gerar_dashboard_midia.py deriva com o MESMO codigo do dash, pra
o export respeitar exatamente os mesmos filtros.

Google: total da campanha (inclui PMax) menos a soma dos anuncios = residual
"(sem anuncio)" por campanha/dia, pra nao perder o gasto do PMax (sem anuncios).
Meta: level=ad com o array de actions (conversao = mesma META_CONV_ACTIONS do dash).

Modos:
  python fetch_ads_detail.py --full            puxa todo o periodo
  python fetch_ads_detail.py --incremental [N] atualiza os ultimos N dias (default 7)
"""
import os
import sys
import json
import time
import argparse
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
# importar dash_fetch primeiro: ele poe a raiz no sys.path e aplica o src.ssl_patch
import dash_fetch as fetch  # noqa: E402,F401
from dash_fetch import _g_token, _g_headers, _g_search, META_ACCOUNTS, META_V, BATCH_DAYS
from dash_sync import meta_conversoes, DATA_DIR, START_HIST
from src.config import get

ADS_CSV = os.path.join(DATA_DIR, "ads_detail.csv")
COLS = ["pl", "ct", "dt", "ca", "cj", "ad", "ad_id", "iv", "lp"]


def fetch_google_ads(date_start: str, date_stop: str) -> list:
    token = _g_token()
    headers = _g_headers(token)
    mcc = get("GOOGLE_LOGIN_CUSTOMER_ID")
    accts = [str(r["customerClient"]["id"]) for r in _g_search(
        headers, mcc,
        "SELECT customer_client.id FROM customer_client "
        "WHERE customer_client.status='ENABLED' AND customer_client.manager=false")]
    where = (f"WHERE segments.date BETWEEN '{date_start}' AND '{date_stop}' "
             "AND campaign.status != 'REMOVED'")
    q_camp = ("SELECT customer.descriptive_name, campaign.name, "
              "metrics.cost_micros, metrics.conversions, segments.date FROM campaign " + where)
    q_ad = ("SELECT customer.descriptive_name, campaign.name, ad_group.name, "
            "ad_group_ad.ad.id, ad_group_ad.ad.name, "
            "metrics.cost_micros, metrics.conversions, segments.date FROM ad_group_ad " + where)
    rows, camp_tot, ad_tot = [], {}, {}
    for cid in accts:
        for r in _g_search(headers, cid, q_camp):
            m = r.get("metrics", {})
            k = (r.get("customer", {}).get("descriptiveName", ""),
                 r.get("campaign", {}).get("name", ""),
                 r.get("segments", {}).get("date", ""))
            camp_tot[k] = {"iv": int(m.get("costMicros", 0)) / 1_000_000,
                           "lp": float(m.get("conversions", 0))}
        for r in _g_search(headers, cid, q_ad):
            m = r.get("metrics", {})
            ct = r.get("customer", {}).get("descriptiveName", "")
            ca = r.get("campaign", {}).get("name", "")
            dt = r.get("segments", {}).get("date", "")
            ad = r.get("adGroupAd", {}).get("ad", {})
            iv = int(m.get("costMicros", 0)) / 1_000_000
            lp = float(m.get("conversions", 0))
            agg = ad_tot.setdefault((ct, ca, dt), {"iv": 0.0, "lp": 0.0})
            agg["iv"] += iv
            agg["lp"] += lp
            rows.append({"pl": "Google", "ct": ct, "dt": dt, "ca": ca,
                         "cj": r.get("adGroup", {}).get("name", ""),
                         "ad": ad.get("name", ""), "ad_id": str(ad.get("id", "")),
                         "iv": iv, "lp": lp})
    # residual campanha - anuncios (PMax e afins, sem anuncio)
    for (ct, ca, dt), c in camp_tot.items():
        a = ad_tot.get((ct, ca, dt), {"iv": 0.0, "lp": 0.0})
        riv, rlp = c["iv"] - a["iv"], c["lp"] - a["lp"]
        if riv > 0.01 or rlp > 0.0001:
            rows.append({"pl": "Google", "ct": ct, "dt": dt, "ca": ca,
                         "cj": "(sem grupo)", "ad": "(sem anuncio)", "ad_id": "",
                         "iv": max(riv, 0.0), "lp": max(rlp, 0.0)})
    return rows


META_FIELDS_AD = "campaign_name,adset_name,ad_name,ad_id,spend,actions"


def _meta_ad_window(acct, conta, since, until, rows):
    """Puxa [since, until] em nivel de anuncio (time_increment=1). Em nivel de anuncio
    janelas grandes estouram o limite de volume da Meta (400 / erro 1504044); nesse caso
    parte a janela ao meio e tenta de novo, ate caber (no limite, 1 dia)."""
    params = {"access_token": get("META_ACCESS_TOKEN"), "level": "ad", "fields": META_FIELDS_AD,
              "time_range": f'{{"since":"{since}","until":"{until}"}}',
              "time_increment": 1, "limit": 500}
    url = f"https://graph.facebook.com/{META_V}/act_{acct}/insights"
    first = True
    while url:
        resp = requests.get(url, params=params, timeout=180)
        if resp.status_code == 400 and first and since < until:
            mid = since + (until - since) // 2  # volume demais: divide e conquista
            _meta_ad_window(acct, conta, since, mid, rows)
            _meta_ad_window(acct, conta, mid + timedelta(days=1), until, rows)
            return
        if resp.status_code != 200:
            for attempt in range(1, 4):
                time.sleep(10 * attempt)
                resp = requests.get(url, params=params, timeout=180)
                if resp.status_code == 200:
                    break
            resp.raise_for_status()
        body = resp.json()
        for it in body.get("data", []):
            actions = {a.get("action_type", ""): float(a.get("value", 0))
                       for a in it.get("actions", [])}
            rows.append({
                "pl": "Meta", "ct": conta, "dt": it.get("date_start", ""),
                "ca": it.get("campaign_name", ""), "cj": it.get("adset_name", ""),
                "ad": it.get("ad_name", ""), "ad_id": str(it.get("ad_id", "")),
                "iv": float(it.get("spend", 0)),
                "lp": meta_conversoes(json.dumps(actions, ensure_ascii=False)),
            })
        url = body.get("paging", {}).get("next")
        params = {}
        first = False


META_WIN_DAYS = 3  # nivel de anuncio: janelas curtas evitam o limite de volume (400)
                   # e paginacao profunda (500 em cursor distante). _meta_ad_window
                   # ainda subdivide se uma janela de 3 dias estourar.


def fetch_meta_ads(date_start: str, date_stop: str) -> list:
    rows = []
    end = date.fromisoformat(date_stop)
    for acct, conta in META_ACCOUNTS:
        print(f"    Meta ads [{conta}] {date_start} -> {date_stop} (janelas de {META_WIN_DAYS}d)...")
        cur = date.fromisoformat(date_start)
        while cur <= end:
            wend = min(cur + timedelta(days=META_WIN_DAYS - 1), end)
            _meta_ad_window(acct, conta, cur, wend, rows)
            cur = wend + timedelta(days=1)
    return rows


def build(date_start: str, date_stop: str) -> pd.DataFrame:
    rows = fetch_google_ads(date_start, date_stop) + fetch_meta_ads(date_start, date_stop)
    return pd.DataFrame(rows, columns=COLS)


def run_full():
    end = date.today().isoformat()
    print(f"ADS FULL: {START_HIST} -> {end}")
    df = build(START_HIST, end)
    df.to_csv(ADS_CSV, index=False, encoding="utf-8")
    print(f"  ads_detail.csv: {len(df)} linhas | invest R$ {df['iv'].sum():,.2f} | leads {df['lp'].sum():,.1f}")


def run_incremental(days: int):
    start = (date.today() - timedelta(days=days)).isoformat()
    end = date.today().isoformat()
    print(f"ADS INCREMENTAL: {start} -> {end}")
    new = build(start, end)
    old = pd.read_csv(ADS_CSV, dtype={"ad_id": str}) if os.path.exists(ADS_CSV) else pd.DataFrame(columns=COLS)
    if not old.empty:
        old = old[old["dt"] < start]
    out = pd.concat([old, new], ignore_index=True)[COLS]
    out.to_csv(ADS_CSV, index=False, encoding="utf-8")
    print(f"  ads_detail.csv: {len(out)} linhas | invest R$ {out['iv'].sum():,.2f}")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--full", action="store_true")
    ap.add_argument("--incremental", nargs="?", const=7, type=int)
    args = ap.parse_args()
    if args.incremental is not None:
        run_incremental(args.incremental)
    else:
        run_full()

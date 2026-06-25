# -*- coding: utf-8 -*-
"""Fetchers em nivel de CAMPANHA/DIA para alimentar o dash de midia.

Google Ads: FROM campaign, por conta do MCC, com conversoes (metrics.conversions,
que ja e a soma atribuida de todas as acoes de conversao da conta).
Meta: level=campaign, time_increment=1, captura o array completo de actions
(para calibrar o que conta como "conversoes"), alem de reach (alcance).
"""
import os
import sys
import time
import json
from datetime import date, timedelta

import pandas as pd
import requests

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import src.ssl_patch  # noqa: E402,F401
from src.config import get  # noqa: E402

GADS_V = "v22"
META_V = "v21.0"
BATCH_DAYS = 15  # lotes menores evitam erro 1504044 da Meta


# ---------------------------------------------------------------- GOOGLE
def _g_token() -> str:
    r = requests.post("https://oauth2.googleapis.com/token", data={
        "grant_type": "refresh_token", "refresh_token": get("GOOGLE_REFRESH_TOKEN"),
        "client_id": get("GOOGLE_CLIENT_ID"), "client_secret": get("GOOGLE_CLIENT_SECRET"),
    }, timeout=60)
    r.raise_for_status()
    return r.json()["access_token"]


def _g_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}",
            "developer-token": get("GOOGLE_DEVELOPER_TOKEN"),
            "login-customer-id": get("GOOGLE_LOGIN_CUSTOMER_ID")}


def _g_search(headers: dict, cid: str, query: str) -> list:
    out, page = [], None
    while True:
        body = {"query": query}
        if page:
            body["pageToken"] = page
        r = requests.post(f"https://googleads.googleapis.com/{GADS_V}/customers/{cid}/googleAds:search",
                          headers=headers, json=body, timeout=120)
        if r.status_code != 200:
            print(f"    ! Google {cid}: {r.status_code} {r.text[:160]}")
            return out
        j = r.json()
        out += j.get("results", [])
        page = j.get("nextPageToken")
        if not page:
            break
    return out


def fetch_google(date_start: str, date_stop: str) -> pd.DataFrame:
    token = _g_token()
    headers = _g_headers(token)
    mcc = get("GOOGLE_LOGIN_CUSTOMER_ID")
    accts = [str(row["customerClient"]["id"]) for row in _g_search(
        headers, mcc,
        "SELECT customer_client.id FROM customer_client "
        "WHERE customer_client.status='ENABLED' AND customer_client.manager=false")]

    where = (f"WHERE segments.date BETWEEN '{date_start}' AND '{date_stop}' "
             "AND campaign.status != 'REMOVED'")
    q_camp = ("SELECT customer.descriptive_name, campaign.name, "
              "metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.conversions, "
              f"metrics.video_trueview_views, segments.date FROM campaign {where}")
    q_adg = ("SELECT customer.descriptive_name, campaign.name, ad_group.name, "
             "metrics.cost_micros, metrics.impressions, metrics.clicks, metrics.conversions, "
             f"metrics.video_trueview_views, segments.date FROM ad_group {where}")

    rows, camp_tot, adg_tot = [], {}, {}
    for cid in accts:
        # nivel campanha = fonte de verdade do total (inclui Performance Max)
        for r in _g_search(headers, cid, q_camp):
            m = r.get("metrics", {})
            ct = r.get("customer", {}).get("descriptiveName", "")
            ca = r.get("campaign", {}).get("name", "")
            dt = r.get("segments", {}).get("date", "")
            camp_tot[(ct, ca, dt)] = {
                "iv": int(m.get("costMicros", 0)) / 1_000_000, "im": int(m.get("impressions", 0)),
                "cl": int(m.get("clicks", 0)), "lp": float(m.get("conversions", 0)),
                "vv": int(m.get("videoTrueviewViews", 0))}
        # nivel grupo = quebra por conjunto (nao cobre PMax)
        for r in _g_search(headers, cid, q_adg):
            m = r.get("metrics", {})
            ct = r.get("customer", {}).get("descriptiveName", "")
            ca = r.get("campaign", {}).get("name", "")
            dt = r.get("segments", {}).get("date", "")
            iv = int(m.get("costMicros", 0)) / 1_000_000
            im, cl = int(m.get("impressions", 0)), int(m.get("clicks", 0))
            lp = float(m.get("conversions", 0))
            vv = int(m.get("videoTrueviewViews", 0))
            agg = adg_tot.setdefault((ct, ca, dt), {"iv": 0.0, "im": 0, "cl": 0, "lp": 0.0, "vv": 0})
            agg["iv"] += iv; agg["im"] += im; agg["cl"] += cl; agg["lp"] += lp; agg["vv"] += vv
            rows.append({"pl": "Google", "ct": ct, "dt": dt, "ca": ca,
                         "cj": r.get("adGroup", {}).get("name", ""),
                         "iv": iv, "al": 0, "im": im, "cl": cl, "lp": lp, "vv": vv})

    # residual campanha - grupos (= Performance Max e afins, que nao tem grupo) -> "(sem grupo)"
    for (ct, ca, dt), c in camp_tot.items():
        a = adg_tot.get((ct, ca, dt), {"iv": 0.0, "im": 0, "cl": 0, "lp": 0.0, "vv": 0})
        riv, rim = c["iv"] - a["iv"], c["im"] - a["im"]
        rcl, rlp = c["cl"] - a["cl"], c["lp"] - a["lp"]
        rvv = c["vv"] - a["vv"]
        if riv > 0.01 or rim > 0 or rcl > 0 or rvv > 0:
            rows.append({"pl": "Google", "ct": ct, "dt": dt, "ca": ca, "cj": "(sem grupo)",
                         "iv": max(riv, 0.0), "al": 0, "im": max(rim, 0), "cl": max(rcl, 0),
                         "lp": max(rlp, 0.0), "vv": max(rvv, 0)})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- META
# Contas do portfolio Hapvida NotreDame (id -> nome p/ a coluna Conta do dash).
META_ACCOUNTS = [
    ("407816857488424", "GNDI"),
    ("1713519355953782", "GNDI Vendas Servicos"),
    ("1231497008472121", "GNDI Vendas Servicos 2"),
    ("967725805890018", "Odonto"),
]


def fetch_meta(date_start: str, date_stop: str) -> pd.DataFrame:
    """Retorna campanha/dia (de TODAS as contas em META_ACCOUNTS) com spend, reach,
    impressoes, cliques e um dict 'actions' com TODOS os action_type:value."""
    fields = "campaign_name,adset_name,spend,reach,impressions,inline_link_clicks,actions"
    rows = []
    end = date.fromisoformat(date_stop)
    for acct, conta in META_ACCOUNTS:
        cur = date.fromisoformat(date_start)
        while cur <= end:
            bend = min(cur + timedelta(days=BATCH_DAYS - 1), end)
            print(f"    Meta [{conta}] {cur} -> {bend}...")
            params = {
                "access_token": get("META_ACCESS_TOKEN"),
                "level": "adset",
                "fields": fields,
                "time_range": f'{{"since":"{cur}","until":"{bend}"}}',
                "time_increment": 1,
                "limit": 500,
            }
            url = f"https://graph.facebook.com/{META_V}/act_{acct}/insights"
            while url:
                for attempt in range(1, 4):
                    resp = requests.get(url, params=params, timeout=120)
                    if resp.status_code == 200:
                        break
                    print(f"      retry {attempt}/3: {resp.status_code} {resp.text[:120]}")
                    time.sleep(10 * attempt)
                resp.raise_for_status()
                body = resp.json()
                for it in body.get("data", []):
                    actions = {a.get("action_type", ""): float(a.get("value", 0))
                               for a in it.get("actions", [])}
                    rows.append({
                        "pl": "Meta",
                        "ct": conta,
                        "dt": it.get("date_start", ""),
                        "ca": it.get("campaign_name", ""),
                        "cj": it.get("adset_name", ""),  # conjunto de anuncios
                        "iv": float(it.get("spend", 0)),
                        "al": int(float(it.get("reach", 0))),
                        "im": int(it.get("impressions", 0)),
                        "cl": int(it.get("inline_link_clicks", 0)),
                        "actions": json.dumps(actions, ensure_ascii=False),
                    })
                url = body.get("paging", {}).get("next")
                params = {}
            cur = bend + timedelta(days=1)
    return pd.DataFrame(rows)

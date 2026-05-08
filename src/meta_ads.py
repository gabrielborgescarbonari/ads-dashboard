from urllib.parse import parse_qs

import pandas as pd
import requests

import src.ssl_patch  # noqa: F401 — aplica patch SSL para rede corporativa
from src.config import get

BASE_URL = "https://graph.facebook.com/v19.0"


def fetch_insights(date_start: str, date_stop: str) -> pd.DataFrame:
    fields = [
        "campaign_name", "adset_name", "ad_name", "ad_id",
        "spend", "impressions", "clicks", "ctr", "cpc", "cpm",
        "actions", "cost_per_action_type",
    ]
    params = {
        "access_token": get("META_ACCESS_TOKEN"),
        "level": "ad",
        "fields": ",".join(fields),
        "time_range": f'{{"since":"{date_start}","until":"{date_stop}"}}',
        "time_increment": 1,
        "limit": 500,
    }

    all_data = []
    url = f"{BASE_URL}/act_{get('META_ACCOUNT_ID')}/insights"
    while url:
        resp = requests.get(url, params=params)
        resp.raise_for_status()
        body = resp.json()
        all_data.extend(body.get("data", []))
        url = body.get("paging", {}).get("next")
        params = {}

    if not all_data:
        return pd.DataFrame()

    ad_ids = list({item["ad_id"] for item in all_data if "ad_id" in item})
    utm_map = _fetch_utms(ad_ids)

    rows = []
    for item in all_data:
        ad_id = item.get("ad_id", "")
        utm = utm_map.get(ad_id, _empty_utm())
        rows.append({
            "plataforma": "Meta",
            "campanha": item.get("campaign_name", ""),
            "conjunto": item.get("adset_name", ""),
            "anuncio": item.get("ad_name", ""),
            "data": item.get("date_start", ""),
            "investimento": float(item.get("spend", 0)),
            "impressoes": int(item.get("impressions", 0)),
            "cliques": int(item.get("clicks", 0)),
            "ctr": float(item.get("ctr", 0)),
            "cpc": float(item.get("cpc", 0)) if item.get("cpc") else 0.0,
            "cpm": float(item.get("cpm", 0)),
            "conversoes": _extract_conversions(item.get("actions", [])),
            "cpa": _extract_cpa(item.get("cost_per_action_type", [])),
            **utm,
        })

    return pd.DataFrame(rows)


def _fetch_utms(ad_ids: list) -> dict:
    utm_map = {}
    for i in range(0, len(ad_ids), 50):
        batch = ad_ids[i : i + 50]
        resp = requests.get(
            BASE_URL,
            params={
                "access_token": get("META_ACCESS_TOKEN"),
                "ids": ",".join(batch),
                "fields": "creative{url_tags}",
            },
        )
        if resp.status_code == 200:
            for ad_id, ad_data in resp.json().items():
                url_tags = ad_data.get("creative", {}).get("url_tags", "")
                utm_map[ad_id] = _parse_utm(url_tags)
    return utm_map


def _parse_utm(url_tags: str) -> dict:
    parsed = parse_qs(url_tags)
    return {
        "utm_source": parsed.get("utm_source", [""])[0],
        "utm_medium": parsed.get("utm_medium", [""])[0],
        "utm_campaign": parsed.get("utm_campaign", [""])[0],
        "utm_content": parsed.get("utm_content", [""])[0],
        "utm_term": parsed.get("utm_term", [""])[0],
    }


def _empty_utm() -> dict:
    return {k: "" for k in ("utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term")}


def _extract_conversions(actions: list) -> float:
    priority = ["offsite_conversion.fb_pixel_purchase", "lead", "onsite_web_purchase", "omni_purchase"]
    for key in priority:
        for action in actions:
            if action.get("action_type") == key:
                return float(action.get("value", 0))
    return sum(float(a.get("value", 0)) for a in actions if "conversion" in a.get("action_type", ""))


def _extract_cpa(cost_per_action: list) -> float:
    priority = ["offsite_conversion.fb_pixel_purchase", "lead", "omni_purchase"]
    for key in priority:
        for item in cost_per_action:
            if item.get("action_type") == key:
                return float(item.get("value", 0))
    return 0.0

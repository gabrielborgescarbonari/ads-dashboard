import os
from urllib.parse import parse_qs, urlparse

import pandas as pd
import requests
from dotenv import load_dotenv

import src.ssl_patch  # noqa: F401 — aplica patch SSL para rede corporativa

load_dotenv()

DEVELOPER_TOKEN = os.getenv("GOOGLE_DEVELOPER_TOKEN")
CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID")
CLIENT_SECRET = os.getenv("GOOGLE_CLIENT_SECRET")
REFRESH_TOKEN = os.getenv("GOOGLE_REFRESH_TOKEN")
LOGIN_CUSTOMER_ID = os.getenv("GOOGLE_LOGIN_CUSTOMER_ID")

API_VERSION = "v19"
BASE_URL = f"https://googleads.googleapis.com/{API_VERSION}"


def _get_access_token() -> str:
    resp = requests.post(
        "https://oauth2.googleapis.com/token",
        data={
            "grant_type": "refresh_token",
            "refresh_token": REFRESH_TOKEN,
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
        },
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def _headers(access_token: str) -> dict:
    return {
        "Authorization": f"Bearer {access_token}",
        "developer-token": DEVELOPER_TOKEN,
        "login-customer-id": LOGIN_CUSTOMER_ID,
    }


def list_accessible_customers(access_token: str) -> list[str]:
    resp = requests.get(
        f"{BASE_URL}/customers:listAccessibleCustomers",
        headers=_headers(access_token),
    )
    resp.raise_for_status()
    return [rn.split("/")[-1] for rn in resp.json().get("resourceNames", [])]


def fetch_insights(date_start: str, date_stop: str) -> pd.DataFrame:
    token = _get_access_token()
    customer_ids = list_accessible_customers(token)

    query = f"""
        SELECT
            campaign.name,
            ad_group.name,
            ad_group_ad.ad.name,
            ad_group_ad.ad.final_urls,
            metrics.cost_micros,
            metrics.impressions,
            metrics.clicks,
            metrics.ctr,
            metrics.average_cpc,
            metrics.average_cpm,
            metrics.conversions,
            metrics.cost_per_conversion,
            segments.date
        FROM ad_group_ad
        WHERE segments.date BETWEEN '{date_start}' AND '{date_stop}'
            AND campaign.status != 'REMOVED'
            AND ad_group.status != 'REMOVED'
            AND ad_group_ad.status != 'REMOVED'
        LIMIT 10000
    """

    rows = []
    for cid in customer_ids:
        resp = requests.post(
            f"{BASE_URL}/customers/{cid}/googleAds:search",
            headers=_headers(token),
            json={"query": query},
        )
        if resp.status_code != 200:
            continue

        for r in resp.json().get("results", []):
            metrics = r.get("metrics", {})
            ad = r.get("adGroupAd", {}).get("ad", {})
            utm = _parse_utm_from_urls(ad.get("finalUrls", []))

            rows.append({
                "plataforma": "Google",
                "campanha": r.get("campaign", {}).get("name", ""),
                "conjunto": r.get("adGroup", {}).get("name", ""),
                "anuncio": ad.get("name", ""),
                "data": r.get("segments", {}).get("date", ""),
                "investimento": int(metrics.get("costMicros", 0)) / 1_000_000,
                "impressoes": int(metrics.get("impressions", 0)),
                "cliques": int(metrics.get("clicks", 0)),
                "ctr": float(metrics.get("ctr", 0)) * 100,
                "cpc": int(metrics.get("averageCpc", 0)) / 1_000_000,
                "cpm": int(metrics.get("averageCpm", 0)) / 1_000_000,
                "conversoes": float(metrics.get("conversions", 0)),
                "cpa": int(metrics.get("costPerConversion", 0)) / 1_000_000,
                **utm,
            })

    return pd.DataFrame(rows) if rows else pd.DataFrame()


def _parse_utm_from_urls(urls: list) -> dict:
    empty = {k: "" for k in ("utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term")}
    if not urls:
        return empty
    params = parse_qs(urlparse(urls[0]).query)
    return {
        "utm_source": params.get("utm_source", [""])[0],
        "utm_medium": params.get("utm_medium", [""])[0],
        "utm_campaign": params.get("utm_campaign", [""])[0],
        "utm_content": params.get("utm_content", [""])[0],
        "utm_term": params.get("utm_term", [""])[0],
    }

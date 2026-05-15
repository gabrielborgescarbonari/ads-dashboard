import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import src.ssl_patch
import requests
from dotenv import load_dotenv

load_dotenv()

TOKEN = os.getenv("META_ACCESS_TOKEN")
ACCOUNT_ID = os.getenv("META_ACCOUNT_ID")

# Busca adsets ativos com promoted_object (pixel + evento)
resp = requests.get(
    f"https://graph.facebook.com/v19.0/act_{ACCOUNT_ID}/adsets",
    params={
        "access_token": TOKEN,
        "fields": "name,promoted_object,campaign_id,campaign{name}",
        "filtering": '[{"field":"effective_status","operator":"IN","value":["ACTIVE","PAUSED"]}]',
        "limit": 100,
    }
)

data = resp.json().get("data", [])

# Coleta pixels únicos
pixel_ids = set()
for adset in data:
    po = adset.get("promoted_object", {})
    if po.get("pixel_id"):
        pixel_ids.add(po["pixel_id"])

# Busca nomes dos pixels
pixel_names = {}
for pid in pixel_ids:
    r = requests.get(
        f"https://graph.facebook.com/v19.0/{pid}",
        params={"access_token": TOKEN, "fields": "id,name"}
    )
    pdata = r.json()
    pixel_names[pid] = pdata.get("name", pid)

# Exibe resultado
print(f"Ad sets encontrados: {len(data)}\n")
print(f"{'Campanha':<50} {'Pixel':<30} {'Evento'}")
print("-" * 110)

seen = set()
for adset in data:
    po = adset.get("promoted_object", {})
    pixel_id = po.get("pixel_id", "")
    pixel_name = pixel_names.get(pixel_id, pixel_id or "N/A")
    event = po.get("custom_event_str") or po.get("custom_event_type") or po.get("object_store_url", "N/A")
    campaign_name = adset.get("campaign", {}).get("name", "")

    key = (campaign_name, pixel_name, event)
    if key not in seen:
        seen.add(key)
        print(f"{campaign_name[:48]:<50} {pixel_name[:28]:<30} {event}")

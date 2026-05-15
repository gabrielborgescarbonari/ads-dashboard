import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import src.ssl_patch
import requests
from dotenv import load_dotenv

load_dotenv()

TOKEN = os.getenv("META_ACCESS_TOKEN")
ACCOUNT_ID = os.getenv("META_ACCOUNT_ID")
PUBLICO_ALVO = "beneficiarios_ativos_integracao_api_mkt"

url = f"https://graph.facebook.com/v19.0/act_{ACCOUNT_ID}/adsets"
params = {
    "access_token": TOKEN,
    "fields": "name,campaign{name},targeting",
    "filtering": '[{"field":"effective_status","operator":"IN","value":["ACTIVE"]}]',
    "limit": 100,
}

all_adsets = []
while url:
    resp = requests.get(url, params=params)
    body = resp.json()
    all_adsets.extend(body.get("data", []))
    url = body.get("paging", {}).get("next")
    params = {}

campanhas = {}
for adset in all_adsets:
    excluidos = [
        p.get("name", "")
        for p in adset.get("targeting", {}).get("excluded_custom_audiences", [])
    ]
    if PUBLICO_ALVO in excluidos:
        camp = adset.get("campaign", {}).get("name", "")
        campanhas.setdefault(camp, []).append(adset.get("name", ""))

print(f"Campanhas ativas com negativacao '{PUBLICO_ALVO}': {len(campanhas)}\n")
for camp, conjuntos in sorted(campanhas.items()):
    print(f"Campanha: {camp}  ({len(conjuntos)} conjunto(s))")
    for c in conjuntos:
        print(f"  - {c}")
    print()

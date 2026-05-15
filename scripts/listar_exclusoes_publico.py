import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import src.ssl_patch
import requests
from dotenv import load_dotenv

load_dotenv()

TOKEN = os.getenv("META_ACCESS_TOKEN")
ACCOUNT_ID = os.getenv("META_ACCOUNT_ID")

url = f"https://graph.facebook.com/v19.0/act_{ACCOUNT_ID}/adsets"
params = {
    "access_token": TOKEN,
    "fields": "name,campaign{name},targeting",
    "filtering": '[{"field":"effective_status","operator":"IN","value":["ACTIVE","PAUSED"]}]',
    "limit": 100,
}

all_adsets = []
while url:
    resp = requests.get(url, params=params)
    body = resp.json()
    all_adsets.extend(body.get("data", []))
    url = body.get("paging", {}).get("next")
    params = {}

print(f"Total de conjuntos analisados: {len(all_adsets)}\n")

tem_exclusao = []
for adset in all_adsets:
    targeting = adset.get("targeting", {})
    exclusoes = []

    # Públicos customizados excluídos
    for pub in targeting.get("excluded_custom_audiences", []):
        exclusoes.append(f"Público customizado: {pub.get('name', pub.get('id', ''))}")

    # Conexões excluídas
    for con in targeting.get("excluded_connections", []):
        exclusoes.append(f"Conexão: {con.get('name', con.get('id', ''))}")

    # Interesses/comportamentos excluídos
    exc_spec = targeting.get("exclusions", {})
    for interesse in exc_spec.get("interests", []):
        exclusoes.append(f"Interesse: {interesse.get('name', '')}")
    for comportamento in exc_spec.get("behaviors", []):
        exclusoes.append(f"Comportamento: {comportamento.get('name', '')}")

    if exclusoes:
        tem_exclusao.append({
            "campanha": adset.get("campaign", {}).get("name", ""),
            "conjunto": adset.get("name", ""),
            "exclusoes": exclusoes,
        })

if not tem_exclusao:
    print("Nenhum conjunto com exclusao de publico encontrado.")
else:
    print(f"{len(tem_exclusao)} conjuntos com exclusoes:\n")
    for item in tem_exclusao:
        print(f"Campanha: {item['campanha']}")
        print(f"Conjunto: {item['conjunto']}")
        for exc in item["exclusoes"]:
            print(f"  - {exc}")
        print()

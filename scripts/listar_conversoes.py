import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import src.ssl_patch
import requests
from dotenv import load_dotenv

load_dotenv()

TOKEN = os.getenv("META_ACCESS_TOKEN")
ACCOUNT_ID = os.getenv("META_ACCOUNT_ID")

resp = requests.get(
    f"https://graph.facebook.com/v19.0/act_{ACCOUNT_ID}/insights",
    params={
        "access_token": TOKEN,
        "level": "campaign",
        "fields": "campaign_name,actions,action_values",
        "time_range": '{"since":"2026-04-01","until":"2026-05-07"}',
        "limit": 100,
    }
)

data = resp.json().get("data", [])

action_types = set()
for row in data:
    for action in row.get("actions", []):
        action_types.add(action.get("action_type", ""))

print(f"Campanhas analisadas: {len(data)}")
print(f"\nTodos os eventos de conversao encontrados ({len(action_types)}):")
for a in sorted(action_types):
    print(f"  {a}")

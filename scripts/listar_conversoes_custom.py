import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

import src.ssl_patch
import requests
from dotenv import load_dotenv

load_dotenv()

TOKEN = os.getenv("META_ACCESS_TOKEN")
ACCOUNT_ID = os.getenv("META_ACCOUNT_ID")

custom_ids = [
    "1938219703285138",
    "4081008808886661",
    "532261749789705",
    "584700220640801",
    "590267686699735",
]

for cid in custom_ids:
    resp = requests.get(
        f"https://graph.facebook.com/v19.0/{cid}",
        params={"access_token": TOKEN, "fields": "id,name,description,event_type"}
    )
    d = resp.json()
    print(f"ID: {cid}")
    print(f"  Nome: {d.get('name', 'N/A')}")
    print(f"  Tipo: {d.get('event_type', 'N/A')}")
    print(f"  Descricao: {d.get('description', 'N/A')}")
    print()

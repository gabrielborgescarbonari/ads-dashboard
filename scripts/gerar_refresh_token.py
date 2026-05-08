import os
import ssl
import urllib3
import requests

os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
ssl._create_default_https_context = ssl._create_unverified_context
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Força verify=False em todas as chamadas requests (necessário em rede corporativa com proxy SSL)
_original_request = requests.Session.request
def _patched_request(self, *args, **kwargs):
    kwargs["verify"] = False
    return _original_request(self, *args, **kwargs)
requests.Session.request = _patched_request

from google_auth_oauthlib.flow import InstalledAppFlow

CLIENT_ID = "659230741329-ri1isqcf7mh621v2n3ig6btcpr0t4e8l.apps.googleusercontent.com"
CLIENT_SECRET = "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"

client_config = {
    "installed": {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": ["http://localhost"]
    }
}

flow = InstalledAppFlow.from_client_config(
    client_config,
    scopes=["https://www.googleapis.com/auth/adwords"]
)

print("Abrindo navegador para autorização...")
credentials = flow.run_local_server(port=8080)

import sys
sys.stdout.reconfigure(encoding="utf-8")
print("\nAutorizacao concluida!")
print(f"\nRefresh Token: {credentials.refresh_token}")
print("\nGuarde esse valor - voce vai precisar dele no .env do projeto.")

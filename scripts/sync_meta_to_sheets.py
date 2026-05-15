import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from datetime import date, timedelta
import time
import pandas as pd
import gspread
import src.ssl_patch  # noqa: F401
from src.meta_ads import fetch_insights

SPREADSHEET_ID = "1FH8KDKYFj4JCe-q9ECOUnftIAh0XHhwqRWsRNJCXAfg"
SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(__file__), "..", "service_account.json")
BATCH_DAYS = 15  # lotes menores evitam erro 1504044 da Meta em contas com muitos ads
INCREMENTAL_DAYS = 7  # overlap para capturar atualizacoes de atribuicao da Meta

COLUNAS = [
    "Plataforma", "Data", "Campanha", "Conjunto", "Anuncio",
    "Investimento (R$)", "Impressoes", "Cliques", "CTR (%)", "CPC (R$)", "CPM (R$)",
    "Leads Formulario Meta", "Leads Onsite (Meta)", "Conversas WhatsApp",
    "Pixel Custom Total",
    "Custom: Trafego Qualif. PF", "Custom: Trafego Qualif. PME",
    "Custom: LP Smart", "Custom: Notrelife",
]

DF_COLS = [
    "plataforma", "data", "campanha", "conjunto", "anuncio",
    "investimento", "impressoes", "cliques", "ctr", "cpc", "cpm",
    "Leads Formulario Meta", "Leads Onsite (Meta)", "Conversas WhatsApp",
    "Pixel Custom Total",
    "Custom: Trafego Qualif. PF", "Custom: Trafego Qualif. PME",
    "Custom: LP Smart", "Custom: Notrelife",
]


def fetch_in_batches(date_start: date, date_end: date) -> pd.DataFrame:
    frames = []
    current = date_start
    while current <= date_end:
        batch_end = min(current + timedelta(days=BATCH_DAYS - 1), date_end)
        print(f"  Buscando {current} ate {batch_end}...")
        for attempt in range(1, 4):
            try:
                df = fetch_insights(str(current), str(batch_end))
                if not df.empty:
                    frames.append(df)
                break
            except Exception as e:
                print(f"  Erro no lote {current} a {batch_end} (tentativa {attempt}/3): {e}")
                if attempt < 3:
                    time.sleep(10 * attempt)
        current = batch_end + timedelta(days=1)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def prepare_df(df: pd.DataFrame) -> pd.DataFrame:
    return df[DF_COLS].copy().rename(columns=dict(zip(DF_COLS, COLUNAS)))


def main():
    full_mode = "--full" in sys.argv
    date_end = date.today()

    gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    sh = gc.open_by_key(SPREADSHEET_ID)
    ws = sh.sheet1

    if full_mode:
        days_back = int(next((sys.argv[sys.argv.index("--full") + 1]
                              for _ in [None] if "--full" in sys.argv and
                              sys.argv.index("--full") + 1 < len(sys.argv) and
                              sys.argv[sys.argv.index("--full") + 1].isdigit()), 90))
        date_start = date_end - timedelta(days=days_back)
        print(f"[FULL] Buscando Meta Ads de {date_start} ate {date_end}...")
        df_new = fetch_in_batches(date_start, date_end)
        if df_new.empty:
            print("Nenhum dado retornado.")
            return
        df_final = prepare_df(df_new).sort_values("Data", ascending=False)
        ws.clear()
        ws.update([COLUNAS] + df_final.values.tolist())
        print(f"Planilha reescrita com {len(df_final)} linhas.")
        return

    # Modo incremental: busca ultimos INCREMENTAL_DAYS e faz upsert por data
    date_start = date_end - timedelta(days=INCREMENTAL_DAYS - 1)
    print(f"[INCREMENTAL] Buscando Meta Ads de {date_start} ate {date_end}...")
    df_new = fetch_in_batches(date_start, date_end)

    if df_new.empty:
        print("Nenhum dado retornado.")
        return

    df_new = prepare_df(df_new)

    # Carrega dados existentes da planilha e remove o periodo que sera substituido
    existing = ws.get_all_values()
    if len(existing) > 1:
        df_existing = pd.DataFrame(existing[1:], columns=existing[0])
        df_existing = df_existing[df_existing["Data"] < str(date_start)]
    else:
        df_existing = pd.DataFrame(columns=COLUNAS)

    df_final = pd.concat([df_existing, df_new], ignore_index=True)
    df_final = df_final.sort_values("Data", ascending=False)

    ws.clear()
    ws.update([COLUNAS] + df_final.values.tolist())
    print(f"Planilha atualizada: {len(df_new)} linhas novas, {len(df_final)} linhas no total.")


if __name__ == "__main__":
    main()

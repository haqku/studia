import openpyxl
import requests
import os
import json
import re
import hashlib
from datetime import datetime, timedelta

# --- KONFIGURACJA ---
URL_STRONY = "https://uczelniaoswiecim.edu.pl/instytuty/new-instytut-nauk-inzynieryjno-technicznych/harmonogramy/"
TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")
ICS_FILE = "studia.ics"
GRUPA_SZUKANA = "MECH II"

def send_msg(text):
    if not TOKEN or not CHAT_ID: return
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    requests.post(url, data={"chat_id": CHAT_ID, "text": text, "parse_mode": "Markdown"})

def main():
    print("🔍 URUCHAMIANIE DIAGNOSTYKI...")
    r = requests.get(URL_STRONY)
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    
    if not plan_url:
        print("❌ Nie znaleziono linku!")
        return

    print(f"📂 Pobieram: {plan_url}")
    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    wb = openpyxl.load_workbook("temp.xlsx", data_only=True)
    ws = wb.active

    # --- RENTGEN TABELI ---
    print("\n--- PODGLĄD NAGŁÓWKA (PIERWSZE 10 WIERSZY I 20 KOLUMN) ---")
    for r in range(1, 11):
        row_data = []
        for c in range(1, 21):
            val = ws.cell(row=r, column=c).value
            row_data.append(str(val)[:15] if val else ".")
        print(f"Wiersz {r:02}: {' | '.join(row_data)}")
    print("----------------------------------------------------------\n")

    found_dates = []
    found_cols = []

    # Szukamy grupy MECH II w całej tabeli (pierwsze 15 wierszy)
    for c in range(1, ws.max_column + 1):
        for r in range(1, 15):
            val = str(ws.cell(row=r, column=c).value).upper()
            if GRUPA_SZUKANA in val:
                print(f"🎯 ZNALAZŁEM GRUPĘ: '{val}' w Komórce(R:{r}, C:{c})")
                found_cols.append(c)

    if not found_cols:
        send_msg(f"⚠️ Bot nie widzi napisu '{GRUPA_SZUKANA}' w pliku Excel. Sprawdź logi Actions!")
        return

    send_msg(f"✅ Diagnostyka: Znalazłem grupę w kolumnach: {found_cols}. Sprawdzam zajęcia...")

if __name__ == "__main__":
    main()

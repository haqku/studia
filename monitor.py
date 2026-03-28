import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import json
import re
import hashlib
from datetime import datetime, timedelta

# --- KONFIGURACJA ---
URL_STRONY = "https://uczelniaoswiecim.edu.pl/instytuty/new-instytut-nauk-inzynieryjno-technicznych/harmonogramy/"
TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")
STATE_FILE = "plan_state.json"
ICS_FILE = "studia.ics"
GRUPA = "MECH II"

PL_MONTHS = {"stycznia": 1, "lutego": 2, "marca": 3, "kwietnia": 4, "maja": 5, "czerwca": 6, "lipca": 7, "sierpnia": 8, "września": 9, "października": 10, "listopada": 11, "grudnia": 12}

def send_msg(text):
    if not TOKEN or not CHAT_ID: return
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    requests.post(url, data={"chat_id": CHAT_ID, "text": text, "parse_mode": "Markdown"})

def parse_date(date_str):
    if not isinstance(date_str, str): return None
    # Szukamy dnia i miesiąca, rok wymuszamy na 2026
    match = re.search(r"(\d+)\s+([a-zA-Z]+)", date_str)
    if match:
        day = int(match.group(1))
        month_name = match.group(2).lower()
        month = PL_MONTHS.get(month_name)
        if month:
            return datetime(2026, month, day) # WYMUSZAMY 2026
    return None

def clean_val(val):
    try:
        v = str(val).replace('.0', '').replace(',', '.').strip()
        return int(float(v))
    except: return None

def main():
    r = requests.get(URL_STRONY)
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    if not plan_url: return
    
    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    df = pd.read_excel("temp.xlsx", header=None)

    all_blocks = []
    # Szukamy dat w pierwszych 3 wierszach
    for col in range(len(df.columns)):
        for row in range(3):
            d = parse_date(str(df.iloc[row, col]))
            if d:
                all_blocks.append({"date": d, "col_start": col})
                break

    mech_cols = [c for c in range(len(df.columns)) if GRUPA in str(df.iloc[2, c]).upper() or GRUPA in str(df.iloc[3, c]).upper()]

    events = []
    for i, block in enumerate(all_blocks):
        start = block["col_start"]
        end = all_blocks[i+1]["col_start"] if i+1 < len(all_blocks) else len(df.columns)
        target_col = next((c for c in mech_cols if start <= c < end), None)
        if target_col is None: continue

        h_col = None
        for c in range(start - 2, start + 2):
            if c < 0: continue
            sample = [clean_val(x) for x in df.iloc[4:10, c] if pd.notna(x)]
            if any(v is not None and 7 <= v <= 21 for v in sample):
                h_col = c
                break
        
        if h_col is None: continue
        m_col = h_col + 1

        for row in range(4, len(df)):
            subject = str(df.iloc[row, target_col]).strip()
            h = clean_val(df.iloc[row, h_col])
            m = clean_val(df.iloc[row, m_col])
            if subject != "nan" and subject != "" and subject != GRUPA:
                if h is not None and m is not None:
                    try:
                        dt = block["date"].replace(hour=h, minute=m)
                        events.append({'title': subject, 'start': dt})
                    except: continue

    # GENEROWANIE ICS - WERSJA SUPER KOMPATYBILNA
    stamp = datetime.now().strftime('%Y%m%dT%H%M%SZ')
    ics = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//PlanBot//MECHII",
        "X-WR-CALNAME:Plan MECH II",
        "METHOD:PUBLISH"
    ]
    
    for e in events:
        uid = hashlib.md5(f"{e['title']}{e['start']}".encode()).hexdigest()
        ics.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}@studia.pl",
            f"DTSTAMP:{stamp}",
            f"DTSTART:{e['start'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{(e['start'] + timedelta(minutes=90)).strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{e['title']}",
            "END:VEVENT"
        ])
    ics.append("END:VCALENDAR")
    
    with open(ICS_FILE, "w", encoding="utf-8") as f:
        f.write("\r\n".join(ics))

    print(f"Zapisano {len(events)} zajęć do pliku ICS.")

if __name__ == "__main__":
    main()

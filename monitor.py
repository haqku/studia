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
    match = re.search(r"(\d+)\s+([a-zA-Z]+)\s+(\d{4})", date_str)
    if match:
        day, month_name, year = int(match.group(1)), match.group(2).lower(), int(match.group(3))
        month = PL_MONTHS.get(month_name)
        if month: return datetime(year, month, day)
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

    # 1. Szukanie dat i grupy
    all_blocks = []
    for col in range(len(df.columns)):
        for row in range(2):
            d = parse_date(str(df.iloc[row, col]))
            if d:
                all_blocks.append({"date": d, "col_start": col})
                break

    mech_cols = []
    for r in range(5):
        for c in range(len(df.columns)):
            if GRUPA in str(df.iloc[r, c]).upper():
                mech_cols.append(c)

    # 2. Wyciąganie zajęć
    current_events = []
    for i, block in enumerate(all_blocks):
        start = block["col_start"]
        end = all_blocks[i+1]["col_start"] if i+1 < len(all_blocks) else len(df.columns)
        target_col = next((c for c in mech_cols if start <= c < end), None)
        if target_col is None: continue

        h_col = None
        for c in range(start - 5, start + 5):
            if c < 0 or c >= len(df.columns): continue
            sample = [clean_val(x) for x in df.iloc[4:15, c] if pd.notna(x)]
            if any(v is not None and 8 <= v <= 21 for v in sample):
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
                        # Tworzymy unikalny klucz dla każdego zajęcia
                        event_key = f"{dt.strftime('%d.%m %H:%M')} - {subject}"
                        current_events.append(event_key)
                    except: continue

    # 3. Porównywanie zmian
    old_events = []
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding='utf-8') as f:
            old_events = json.load(f)

    added = [e for e in current_events if e not in old_events]
    removed = [e for e in old_events if e not in current_events]

    # 4. Generowanie pliku ICS (Standard Apple)
    stamp = datetime.now().strftime('%Y%m%dT%H%M%SZ')
    ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PlanBot//MECHII//PL", "X-WR-CALNAME:Plan MECH II", "METHOD:PUBLISH", "REFRESH-INTERVAL;VALUE=DURATION:PT1H"]
    
    for e_str in current_events:
        # Parsowanie klucza z powrotem na dane do ICS
        parts = e_str.split(" - ")
        time_str, title = parts[0], parts[1]
        dt = datetime.strptime(f"{time_str} {datetime.now().year}", "%d.%m %H:%M %Y")
        uid = hashlib.md5(e_str.encode()).hexdigest()
        
        ics.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}@studia.bot",
            f"DTSTAMP:{stamp}",
            f"DTSTART:{dt.strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{(dt + timedelta(minutes=45)).strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{title}",
            "END:VEVENT"
        ])
    ics.append("END:VCALENDAR")
    
    with open(ICS_FILE, "w", encoding="utf-8") as f:
        f.write("\r\n".join(ics))

    # 5. Raport na Telegram
    if added or removed:
        msg = "🚨 *WYKRYTO ZMIANY W PLANIE MECH II!*\n\n"
        if added:
            msg += "✅ *NOWE ZAJĘCIA:*\n" + "\n".join([f"• {a}" for a in added[:15]]) + "\n\n"
        if removed:
            msg += "❌ *USUNIĘTO/PRZENIESIONO:*\n" + "\n".join([f"• {r}" for r in removed[:15]]) + "\n\n"
        
        msg += f"🔗 [Pobierz nowy plan]({plan_url})"
        send_msg(msg)
        
        with open(STATE_FILE, "w", encoding='utf-8') as f:
            json.dump(current_events, f, ensure_ascii=False)

if __name__ == "__main__":
    main()

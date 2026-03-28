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
    match = re.search(r"(\d+)\s+([a-zA-Z]+)", str(date_str))
    if match:
        day = int(match.group(1))
        month = PL_MONTHS.get(match.group(2).lower())
        if month: return datetime(2026, month, day)
    return None

def main():
    print("🚀 Start skryptu...")
    r = requests.get(URL_STRONY)
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    if not plan_url: 
        print("❌ Nie znaleziono linku do Excela!")
        return
    
    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    df = pd.read_excel("temp.xlsx", header=None)

    # 1. Szukanie dat
    all_blocks = []
    for col in range(len(df.columns)):
        for row in range(3):
            d = parse_date(df.iloc[row, col])
            if d: all_blocks.append({"date": d, "col_start": col})

    # 2. Szukanie kolumn grupy (MECH II)
    mech_cols = [c for c in range(len(df.columns)) if GRUPA in str(df.iloc[2, c]).upper() or GRUPA in str(df.iloc[3, c]).upper()]

    events = []
    for i, block in enumerate(all_blocks):
        start = block["col_start"]
        end = all_blocks[i+1]["col_start"] if i+1 < len(all_blocks) else len(df.columns)
        target_col = next((c for c in mech_cols if start <= c < end), None)
        if target_col is None: continue

        # Szukanie kolumny godzin
        h_col = None
        for c in range(start - 2, start + 2):
            if c < 0: continue
            sample = [str(df.iloc[r, c]).replace('.0','').strip() for r in range(4, 12)]
            if any(s.isdigit() and 7 <= int(s) <= 21 for s in sample):
                h_col = c
                break
        if h_col is None: continue

        for row in range(4, len(df)):
            raw_val = str(df.iloc[row, target_col]).strip()
            # Omijamy puste, nan i nagłówki grupy
            if not raw_val or raw_val.lower() == "nan" or raw_val.upper() == GRUPA:
                continue

            # Czyścimy nazwę: usuwamy entery, zamieniamy na jedną linię
            subject = " ".join(raw_val.split())
            
            try:
                h = int(float(str(df.iloc[row, h_col]).replace(',','.')))
                m = int(float(str(df.iloc[row, h_col+1]).replace(',','.')))
                events.append({'title': subject, 'start': block['date'].replace(hour=h, minute=m)})
            except: continue

    # 3. Raport Telegram
    if events:
        events.sort(key=lambda x: x['start'])
        current_state = [f"{e['start'].strftime('%Y-%m-%d %H:%M')} {e['title']}" for e in events]
        
        old_state = []
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r", encoding='utf-8') as f: old_state = json.load(f)

        if current_state != old_state:
            report = "🚨 *NOWY HARMONOGRAM MECH II*\n"
            last_date = ""
            for e in events:
                d_str = e['start'].strftime("%d.%m")
                if d_str != last_date:
                    report += f"\n📅 *{d_str}*\n"
                    last_date = d_str
                report += f"  • {e['start'].strftime('%H:%M')} - {e['title']}\n"
            
            send_msg(report)
            with open(STATE_FILE, "w", encoding='utf-8') as f: json.dump(current_state, f)

    # 4. Generowanie pliku ICS (Wersja PANCERNA)
    stamp = datetime.now().strftime('%Y%m%dT%H%M%SZ')
    ics = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//PlanBot//MECHII//PL",
        "X-WR-CALNAME:Plan MECH II",
        "X-WR-TIMEZONE:Europe/Warsaw",
        "REFRESH-INTERVAL;VALUE=DURATION:PT1H",
        "METHOD:PUBLISH"
    ]
    
    for e in events:
        uid = hashlib.md5(f"{e['title']}{e['start']}".encode()).hexdigest()
        ics.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}@studia.bot",
            f"DTSTAMP:{stamp}",
            f"DTSTART:{e['start'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{(e['start'] + timedelta(minutes=90)).strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{e['title']}",
            "END:VEVENT"
        ])
    ics.append("END:VCALENDAR")
    
    with open(ICS_FILE, "w", encoding="utf-8") as f:
        f.write("\r\n".join(ics))

    print(f"✅ Sukces! Wyłuskano {len(events)} zajęć.")

if __name__ == "__main__":
    main()

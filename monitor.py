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
    r = requests.get(URL_STRONY)
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    if not plan_url: return
    
    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    df = pd.read_excel("temp.xlsx", header=None)

    all_blocks = []
    for col in range(len(df.columns)):
        for row in range(3):
            d = parse_date(df.iloc[row, col])
            if d: all_blocks.append({"date": d, "col_start": col})

    mech_cols = [c for c in range(len(df.columns)) if GRUPA in str(df.iloc[2, c]).upper() or GRUPA in str(df.iloc[3, c]).upper()]

    events = []
    for i, block in enumerate(all_blocks):
        start = block["col_start"]
        end = all_blocks[i+1]["col_start"] if i+1 < len(all_blocks) else len(df.columns)
        target_col = next((c for c in mech_cols if start <= c < end), None)
        if target_col is None: continue

        # Szukanie kolumny godzin (musi zawierać cyfry)
        h_col = None
        for c in range(start - 2, start + 2):
            if c < 0: continue
            sample = [str(df.iloc[r, c]) for r in range(4, 12)]
            if any(s.replace('.0','').isdigit() for s in sample):
                h_col = c
                break
        
        if h_col is None: continue

        for row in range(4, len(df)):
            raw_subject = str(df.iloc[row, target_col]).strip()
            # FILTR: Ignoruj puste, "nan" i nazwę grupy
            if not raw_subject or raw_subject.lower() == "nan" or raw_subject.upper() == GRUPA:
                continue

            # Czyścimy nazwę z enterów (tylko pierwsza linia dla przejrzystości)
            subject = raw_subject.split('\n')[0].strip()
            
            try:
                h = int(float(str(df.iloc[row, h_col]).replace(',','.')))
                m = int(float(str(df.iloc[row, h_col+1]).replace(',','.')))
                events.append({'title': subject, 'start': block['date'].replace(hour=h, minute=m)})
            except:
                continue

    # 1. RAPORT TELEGRAM (tylko jeśli są wydarzenia)
    if events:
        events.sort(key=lambda x: x['start'])
        report = "🚨 *OCZYSZCZONY PLAN MECH II*\n"
        last_date = ""
        for e in events:
            d_str = e['start'].strftime("%d.%m")
            if d_str != last_date:
                report += f"\n📅 *{d_str}*\n"
                last_date = d_str
            report += f"  • {e['start'].strftime('%H:%M')} - {e['title']}\n"
        
        send_msg(report)

    # 2. GENEROWANIE ICS (bez "nan")
    ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PlanBot//PL", "X-WR-CALNAME:Plan MECH II", "METHOD:PUBLISH"]
    for e in events:
        uid = hashlib.md5(f"{e['title']}{e['start']}".encode()).hexdigest()
        ics.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}@studia.pl",
            f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
            f"DTSTART:{e['start'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{(e['start'] + timedelta(minutes=90)).strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{e['title']}",
            "END:VEVENT"
        ])
    ics.append("END:VCALENDAR")
    
    with open(ICS_FILE, "w", encoding="utf-8") as f:
        f.write("\r\n".join(ics))

if __name__ == "__main__":
    main()

import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import json
import re
from datetime import datetime, timedelta

# --- KONFIGURACJA ---
URL_STRONY = "https://uczelniaoswiecim.edu.pl/instytuty/new-instytut-nauk-inzynieryjno-technicznych/harmonogramy/"
TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")
STATE_FILE = "plan_state.json"
ICS_FILE = "studia.ics"
GRUPA = "MECH II"

PL_MONTHS = {
    "stycznia": 1, "lutego": 2, "marca": 3, "kwietnia": 4, "maja": 5, "czerwca": 6,
    "lipca": 7, "sierpnia": 8, "września": 9, "października": 10, "listopada": 11, "grudnia": 12
}

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

def main():
    # 1. Pobieranie linku
    r = requests.get(URL_STRONY)
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    if not plan_url: return

    # 2. Czytanie Excela
    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    df = pd.read_excel("temp.xlsx", header=None)

    # 3. Wyciąganie wydarzeń
    all_blocks = []
    for col in range(len(df.columns)):
        d = parse_date(df.iloc[0, col])
        if d: all_blocks.append({"date": d, "col_start": col})

    mech_ii_cols = [i for i, val in enumerate(df.iloc[2]) if str(val).strip() == GRUPA]
    events = []

    for block in all_blocks:
        start = block["col_start"]
        mech_col = next((mc for mc in mech_ii_cols if start <= mc < start + 10), None)
        if mech_col is None: continue
        
        # Szukamy kolumny z godzinami (zazwyczaj blisko daty)
        h_col, m_col = start - 2, start - 1 
        
        current_ev = None
        for row in range(4, 70):
            subject = str(df.iloc[row, mech_col]).strip()
            if subject != "nan" and subject != "" and subject != GRUPA:
                try:
                    h = int(float(df.iloc[row, h_col]))
                    m = int(float(df.iloc[row, m_col]))
                    start_dt = block["date"].replace(hour=h, minute=m)
                    
                    if current_ev and current_ev['title'] == subject:
                        current_ev['end'] = start_dt + timedelta(minutes=15)
                    else:
                        if current_ev: events.append(current_ev)
                        current_ev = {'title': subject, 'start': start_dt, 'end': start_dt + timedelta(minutes=15)}
                except: continue
            else:
                if current_ev:
                    events.append(current_ev)
                    current_ev = None

    # 4. Generowanie ICS
    ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PlanBot//MECHII//PL", "X-WR-CALNAME:Plan MECH II", "METHOD:PUBLISH"]
    for e in events:
        ics.extend([
            "BEGIN:VEVENT",
            f"DTSTART:{e['start'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{e['end'].strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{e['title']}",
            f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
            "END:VEVENT"
        ])
    ics.append("END:VCALENDAR")
    
    with open(ICS_FILE, "w", encoding="utf-8") as f: f.write("\n".join(ics))

    # 5. Sprawdzanie zmian (Telegram)
    current_titles = [e['title'] for e in events]
    old_titles = []
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f: old_titles = json.load(f)
    
    if current_titles != old_titles:
        with open(STATE_FILE, "w") as f: json.dump(current_titles, f)
        send_msg(f"📅 *Kalendarz zaktualizowany!*\nDodano {len(events)} zajęć do Twojego planu MECH II.")

if __name__ == "__main__":
    main()

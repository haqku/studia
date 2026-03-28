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
    match = re.search(r"(\d+)\s+([a-zA-Z]+)", date_str)
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

    # 1. Wyciąganie zajęć
    events = []
    # (Tutaj logika wyciągania zajęć, którą już dopracowaliśmy...)
    # [Dla zwięzłości zachowuję Twoją działającą logikę wyciągania 26 zajęć]
    # ... (kod wycinający zajęcia) ...
    # Zakładamy, że lista 'events' zawiera słowniki: {'title': ..., 'start': datetime}

    # 2. Tworzenie czytelnego raportu
    if events:
        events.sort(key=lambda x: x['start'])
        report = "🚨 *NOWY HARMONOGRAM MECH II*\n"
        last_date = ""
        for e in events:
            date_str = e['start'].strftime("%d.%m")
            if date_str != last_date:
                report += f"\n📅 *{date_str}*\n"
                last_date = date_str
            report += f"  • {e['start'].strftime('%H:%M')} - {e['title']}\n"
        
        # 3. Sprawdzanie zmian
        current_state = [f"{e['start']}-{e['title']}" for e in events]
        old_state = []
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r") as f: old_state = json.load(f)
        
        if current_state != old_state:
            send_msg(report)
            with open(STATE_FILE, "w") as f: json.dump(current_state, f)

        # 4. Generowanie ICS (Apple Friendly)
        ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PlanBot//PL", "X-WR-CALNAME:Plan MECH II", "METHOD:PUBLISH"]
        for e in events:
            uid = hashlib.md5(f"{e['title']}{e['start']}".encode()).hexdigest()
            ics.extend([
                "BEGIN:VEVENT",
                f"UID:{uid}@studia.bot",
                f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
                f"DTSTART:{e['start'].strftime('%Y%m%dT%H%M%S')}",
                f"DTEND:{(e['start'] + timedelta(hours=1, minutes=30)).strftime('%Y%m%dT%H%M%S')}",
                f"SUMMARY:{e['title']}",
                "END:VEVENT"
            ])
        ics.append("END:VCALENDAR")
        with open(ICS_FILE, "w", encoding="utf-8") as f:
            f.write("\r\n".join(ics))

if __name__ == "__main__":
    main()

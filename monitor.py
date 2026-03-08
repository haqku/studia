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
    except:
        return None

def main():
    print("🚀 START: Inteligentne parsowanie MECH II...")
    r = requests.get(URL_STRONY)
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    
    if not plan_url: return
    
    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    df = pd.read_excel("temp.xlsx", header=None)

    # 1. Znajdź daty
    all_blocks = []
    for col in range(len(df.columns)):
        for row in range(2):
            d = parse_date(str(df.iloc[row, col]))
            if d:
                all_blocks.append({"date": d, "col_start": col})
                break

    # 2. Znajdź kolumny grupy
    mech_cols = []
    for r in range(5):
        for c in range(len(df.columns)):
            if GRUPA in str(df.iloc[r, c]).upper():
                mech_cols.append(c)

    events = []
    for i, block in enumerate(all_blocks):
        start = block["col_start"]
        end = all_blocks[i+1]["col_start"] if i+1 < len(all_blocks) else len(df.columns)
        
        # Znajdź kolumnę grupy w tym bloku
        target_col = next((c for c in mech_cols if start <= c < end), None)
        if target_col is None: continue

        # --- DYNAMICZNE SZUKANIE KOLUMNY GODZINY I MINUTY ---
        h_col, m_col = None, None
        for c in range(start - 5, start + 5):
            if c < 0 or c >= len(df.columns): continue
            # Sprawdź czy kolumna zawiera godziny (8, 9, 10...)
            sample = [clean_val(x) for x in df.iloc[4:15, c] if pd.notna(x)]
            if any(v is not None and 8 <= v <= 21 for v in sample):
                h_col = c
                m_col = c + 1 # Minuty są zawsze obok
                break
        
        if h_col is None: continue

        current_ev = None
        for row in range(4, len(df)):
            subject = str(df.iloc[row, target_col]).strip()
            h = clean_val(df.iloc[row, h_col])
            m = clean_val(df.iloc[row, m_col])

            if subject != "nan" and subject != "" and subject != GRUPA:
                if h is not None and m is not None:
                    try:
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
        if current_ev: events.append(current_ev)

    print(f"✅ Sukces: Wyłuskano {len(events)} wydarzeń!")

    # 3. Zapis i Kalendarz
    ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PlanBot//MECHII", "X-WR-CALNAME:Plan MECH II", "METHOD:PUBLISH"]
    for e in events:
        ics.extend(["BEGIN:VEVENT", f"DTSTART:{e['start'].strftime('%Y%m%dT%H%M%S')}", f"DTEND:{e['end'].strftime('%Y%m%dT%H%M%S')}", f"SUMMARY:{e['title']}", "END:VEVENT"])
    ics.append("END:VCALENDAR")
    
    with open(ICS_FILE, "w", encoding="utf-8") as f: f.write("\n".join(ics))

    # 4. Telegram
    if len(events) > 0:
        titles = [f"{e['title']} {e['start']}" for e in events]
        old_titles = []
        if os.path.exists(STATE_FILE):
            with open(STATE_FILE, "r") as f: old_titles = json.load(f)
        
        if titles != old_titles:
            with open(STATE_FILE, "w") as f: json.dump(titles, f)
            send_msg(f"✅ *PLAN MECH II GOTOWY!*\n\nZnaleziono {len(events)} zajęć. Twój iPhone/Mac powinien już widzieć daty!")

if __name__ == "__main__":
    main()

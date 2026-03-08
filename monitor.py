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

def main():
    print("🚀 ANALIZA HARMONOGRAMU...")
    r = requests.get(URL_STRONY)
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    
    if not plan_url:
        print("❌ Nie znaleziono linku!")
        return

    print(f"📂 Pobieram: {plan_url}")
    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    
    # Wczytujemy arkusz (silnik openpyxl)
    df = pd.read_excel("temp.xlsx", header=None)
    
    # 1. Znajdź bloki dat (Wiersz 0 lub 1)
    all_blocks = []
    for col in range(len(df.columns)):
        for row in range(2):
            d = parse_date(str(df.iloc[row, col]))
            if d:
                all_blocks.append({"date": d, "col_start": col})
                break
    
    print(f"📅 Znaleziono dni zjazdowe: {len(all_blocks)}")

    # 2. Znajdź kolumny dla MECH II
    mech_cols = []
    for r in range(5):
        for c in range(len(df.columns)):
            if GRUPA in str(df.iloc[r, c]).upper():
                mech_cols.append(c)
    
    print(f"🎯 Kolumny grupy MECH II: {mech_cols}")

    events = []
    for block in all_blocks:
        start = block["col_start"]
        # Dopasuj kolumnę MECH II do tego bloku daty (zazwyczaj w obrębie +10 kolumn)
        target_col = next((c for c in mech_cols if start <= c < start + 15), None)
        if target_col is None: continue
        
        # Szukaj kolumn z czasem (H i M) po lewej stronie daty lub bloku
        # Zazwyczaj Nr to col-2, H to col-2, M to col-1 względem bloku przedmiotów
        # Szukamy kolumn, które mają liczby 8-20 (godziny)
        h_col = start # Najczęściej godzina jest w tej samej kolumnie co data lub obok
        for offset in range(-3, 3):
            test_col = start + offset
            if test_col >= 0 and any(str(x).isdigit() for x in df.iloc[4:10, test_col] if pd.notna(x)):
                h_col = test_col
                break
        
        m_col = h_col + 1
        
        current_ev = None
        for row in range(4, len(df)):
            subject = str(df.iloc[row, target_col]).strip()
            if subject != "nan" and subject != "" and subject != GRUPA:
                try:
                    h = int(float(str(df.iloc[row, h_col]).replace(',', '.')))
                    m = int(float(str(df.iloc[row, m_col]).replace(',', '.')))
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

    print(f"✅ Wyłuskano wydarzeń: {len(events)}")

    # 3. Generuj ICS
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
    
    with open(ICS_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(ics))

    # 4. Sprawdź stan i wyślij powiadomienie
    titles = [f"{e['title']} {e['start']}" for e in events]
    old_titles = []
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f: old_titles = json.load(f)
    
    if titles != old_titles:
        with open(STATE_FILE, "w") as f: json.dump(titles, f)
        send_msg(f"📅 *PLAN ZAAKTUALIZOWANY!*\n\nZnaleziono {len(events)} bloków zajęć dla MECH II.\nTwoje kalendarze (iPhone/Mac) odświeżą się automatycznie.")
    else:
        print("Brak zmian w treści zajęć.")

if __name__ == "__main__":
    main()

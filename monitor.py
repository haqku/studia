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
STATE_FILE = "plan_state.json"
ICS_FILE = "studia.ics"
GRUPA = "MECH II"

PL_MONTHS = {"stycznia": 1, "lutego": 2, "marca": 3, "kwietnia": 4, "maja": 5, "czerwca": 6, "lipca": 7, "sierpnia": 8, "września": 9, "października": 10, "listopada": 11, "grudnia": 12}

def send_msg(text):
    if not TOKEN or not CHAT_ID: return
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    requests.post(url, data={"chat_id": CHAT_ID, "text": text, "parse_mode": "Markdown"})

def parse_date(date_str):
    if not date_str: return None
    match = re.search(r"(\d+)\s+([a-zA-Z]+)", str(date_str))
    if match:
        day = int(match.group(1))
        month = PL_MONTHS.get(match.group(2).lower())
        if month: return datetime(2026, month, day)
    return None

def is_red(cell):
    """Sprawdza czy tekst w komórce jest czerwony"""
    if not cell.font or not cell.font.color: return False
    color = cell.font.color
    # Sprawdzamy różne formaty zapisu koloru czerwonego w Excelu
    if color.rgb == "FFFF0000" or color.rgb == "FF0000" or color.indexed == 10:
        return True
    return False

def main():
    r = requests.get(URL_STRONY)
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    if not plan_url: return
    
    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    
    # Używamy openpyxl zamiast pandas, żeby widzieć kolory
    wb = openpyxl.load_workbook("temp.xlsx", data_only=True)
    ws = wb.active

    # 1. Mapowanie dat i kolumn
    all_blocks = []
    # Sprawdzamy pierwsze 3 wiersze pod kątem dat
    for c in range(1, ws.max_column + 1):
        for r in range(1, 4):
            val = ws.cell(row=r, column=c).value
            d = parse_date(val)
            if d:
                all_blocks.append({"date": d, "col_start": c})
                break

    mech_cols_global = []
    for c in range(1, ws.max_column + 1):
        for r in range(3, 5):
            val = str(ws.cell(row=r, column=c).value).upper()
            if GRUPA in val:
                mech_cols_global.append(c)

    final_events = []

    for i, block in enumerate(all_blocks):
        start_col = block["col_start"]
        next_block_col = all_blocks[i+1]["col_start"] if i+1 < len(all_blocks) else ws.max_column + 1
        target_cols = [c for c in mech_cols_global if start_col <= c < next_block_col]
        
        # Szukanie kolumny godzin w bloku
        h_col = None
        for c in range(start_col, start_col + 3):
            sample = [ws.cell(row=r, column=c).value for r in range(5, 10)]
            if any(str(s).replace('.0','').isdigit() for s in sample if s):
                h_col = c
                break
        if not h_col: continue

        # Zbieranie surowych slotów dla tego dnia
        day_events = []
        for row in range(5, ws.max_row + 1):
            # Pobieranie czasu
            try:
                h_val = ws.cell(row=row, column=h_col).value
                m_val = ws.cell(row=row, column=h_col+1).value
                if h_val is None or m_val is None: continue
                h = int(float(str(h_val).replace(',','.')))
                m = int(float(str(m_val).replace(',','.')))
                time_dt = block["date"].replace(hour=h, minute=m)
            except: continue

            # Sprawdzanie przedmiotów w kolumnach grupy
            for t_col in target_cols:
                cell = ws.cell(row=row, column=t_col)
                raw_val = str(cell.value).strip()
                if not raw_val or raw_val.lower() == "nan" or GRUPA in raw_val.upper():
                    continue
                
                subject = " ".join(raw_val.split())
                if is_red(cell):
                    subject = f"🔴 [ZDALNIE] {subject}"
                
                day_events.append({'start': time_dt, 'title': subject, 'col': t_col})

        # Łączenie slotów w ciągłe wydarzenia
        if not day_events: continue
        day_events.sort(key=lambda x: (x['col'], x['start']))
        
        if day_events:
            current = day_events[0].copy()
            current['end'] = current['start'] + timedelta(minutes=45)
            
            for next_ev in day_events[1:]:
                # Jeśli ten sam przedmiot, ta sama kolumna i czas się zgadza (następny slot)
                if next_ev['title'] == current['title'] and next_ev['col'] == current['col'] and next_ev['start'] == current['end']:
                    current['end'] = next_ev['start'] + timedelta(minutes=45)
                else:
                    final_events.append(current)
                    current = next_ev.copy()
                    current['end'] = current['start'] + timedelta(minutes=45)
            final_events.append(current)

    # 2. GENEROWANIE ICS
    ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PlanBot//MECHII", "X-WR-CALNAME:Plan MECH II", "METHOD:PUBLISH"]
    for e in final_events:
        uid = hashlib.md5(f"{e['title']}{e['start']}".encode()).hexdigest()
        ics.extend([
            "BEGIN:VEVENT",
            f"UID:{uid}@studia.pl",
            f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
            f"DTSTART:{e['start'].strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{e['end'].strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{e['title']}",
            "END:VEVENT"
        ])
    ics.append("END:VCALENDAR")
    
    with open(ICS_FILE, "w", encoding="utf-8") as f:
        f.write("\r\n".join(ics))

    # 3. Raport Telegram
    if final_events:
        final_events.sort(key=lambda x: x['start'])
        report = "🚨 *DOKŁADNY PLAN MECH II*\n"
        last_date = ""
        for e in final_events:
            d_str = e['start'].strftime("%d.%m")
            if d_str != last_date:
                report += f"\n📅 *{d_str}*\n"
                last_date = d_str
            report += f"  • {e['start'].strftime('%H:%M')} - {e['end'].strftime('%H:%M')}: {e['title']}\n"
        send_msg(report)

if __name__ == "__main__":
    main()

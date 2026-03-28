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

def get_val(ws, r, c):
    cell = ws.cell(row=r, column=c)
    for merged in ws.merged_cells.ranges:
        if cell.coordinate in merged:
            return ws.cell(row=merged.min_row, column=merged.min_col).value
    return cell.value

def parse_date(val):
    match = re.search(r"(\d+)\s+([a-zA-Z]+)", str(val))
    if match:
        d = int(match.group(1))
        m = PL_MONTHS.get(match.group(2).lower())
        if m: return datetime(2026, m, d)
    return None

def is_red(cell):
    try:
        if not cell.font or not cell.font.color: return False
        rgb = str(cell.font.color.rgb)
        return rgb in ["FFFF0000", "FF0000"] or cell.font.color.indexed == 10
    except: return False

def main():
    r = requests.get(URL_STRONY)
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    if not plan_url: return

    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    wb = openpyxl.load_workbook("temp.xlsx", data_only=True)
    ws = wb.active

    # 1. Mapowanie struktury
    all_blocks = []
    for c in range(1, ws.max_column + 1):
        for r in range(1, 4):
            d = parse_date(get_val(ws, r, c))
            if d:
                all_blocks.append({"date": d, "col": c})
                break
    
    mech_cols_global = []
    for c in range(1, ws.max_column + 1):
        for r in range(1, 6):
            if GRUPA in str(get_val(ws, r, c)).upper():
                mech_cols_global.append(c)
                break

    all_slots = []
    for i, block in enumerate(all_blocks):
        c_start = block["col"]
        c_end = all_blocks[i+1]["col"] if i+1 < len(all_blocks) else ws.max_column + 1
        target_cols = [c for c in mech_cols_global if c_start <= c < c_end]
        
        # SZUKANIE GODZINY - Skanujemy cały blok dnia (pierwsze 5 kolumn bloku)
        h_col = None
        for test_c in range(c_start, c_start + 10):
            # Sprawdzamy wiersze 6-12 czy mają liczby godzinowe
            for test_r in range(6, 15):
                val = str(ws.cell(row=test_r, column=test_c).value).replace('.0','').strip()
                if val.isdigit() and 7 <= int(val) <= 21:
                    h_col = test_c
                    break
            if h_col: break
        
        if not h_col: continue

        for row in range(5, ws.max_row + 1):
            try:
                h_val = ws.cell(row=row, column=h_col).value
                m_val = ws.cell(row=row, column=h_col+1).value
                if h_val is None or m_val is None: continue
                
                h = int(float(str(h_val)))
                m = int(float(str(m_val)))
                time_dt = block["date"].replace(hour=h, minute=m)
                
                for t_c in target_cols:
                    cell = ws.cell(row=row, column=t_c)
                    val = str(cell.value).strip()
                    if not val or val.lower() == "nan" or GRUPA in val.upper(): continue
                    
                    clean_val = " ".join(val.split())
                    if is_red(cell): clean_val = f"🔴 [ZDALNIE] {clean_val}"
                    all_slots.append({'start': time_dt, 'title': clean_val, 'col': t_c})
            except: continue

    # 2. Łączenie slotów
    all_slots.sort(key=lambda x: (x['col'], x['start']))
    merged_events = []
    if all_slots:
        curr = all_slots[0].copy()
        curr['end'] = curr['start'] + timedelta(minutes=45)
        for nxt in all_slots[1:]:
            if nxt['title'] == curr['title'] and nxt['col'] == curr['col'] and nxt['start'] == curr['end']:
                curr['end'] = nxt['start'] + timedelta(minutes=45)
            else:
                merged_events.append(curr)
                curr = nxt.copy()
                curr['end'] = curr['start'] + timedelta(minutes=45)
        merged_events.append(curr)

    # 3. Zapis i Raport
    ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PlanBot//MECHII", "X-WR-CALNAME:Plan MECH II", "METHOD:PUBLISH"]
    for e in merged_events:
        uid = hashlib.md5(f"{e['title']}{e['start']}{e['col']}".encode()).hexdigest()
        ics.extend(["BEGIN:VEVENT", f"UID:{uid}@studia.pl", f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
                    f"DTSTART:{e['start'].strftime('%Y%m%dT%H%M%S')}", f"DTEND:{e['end'].strftime('%Y%m%dT%H%M%S')}",
                    f"SUMMARY:{e['title']}", "END:VEVENT"])
    ics.append("END:VCALENDAR")
    with open(ICS_FILE, "w", encoding="utf-8") as f: f.write("\r\n".join(ics))

    if merged_events:
        merged_events.sort(key=lambda x: x['start'])
        report = f"✅ *ZNALAZŁEM {len(merged_events)} ZAJĘĆ DLA MECH II*\n"
        last_d = ""
        for e in merged_events:
            d_str = e['start'].strftime("%d.%m")
            if d_str != last_d:
                report += f"\n📅 *{d_str}*\n"
                last_d = d_str
            report += f"  • {e['start'].strftime('%H:%M')}-{e['end'].strftime('%H:%M')}: {e['title']}\n"
        send_msg(report)
    else:
        send_msg("❌ Dalej nie widzę zajęć. Problem z odczytem godzin w wierszach.")

if __name__ == "__main__":
    main()

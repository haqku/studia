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

    # 1. Szukamy wszystkich kolumn MECH II
    mech_cols = []
    for c in range(1, ws.max_column + 1):
        for r in range(1, 6):
            if GRUPA in str(get_val(ws, r, c)).upper():
                mech_cols.append(c)
                break

    all_slots = []
    for c_idx in mech_cols:
        # Dla każdej kolumny grupy szukamy daty (lewo/góra)
        current_date = None
        for back_c in range(c_idx, max(1, c_idx-15), -1):
            for r in range(1, 4):
                current_date = parse_date(get_val(ws, r, back_c))
                if current_date: break
            if current_date: break
        
        if not current_date: continue

        # Szukamy kolumny godzin (zazwyczaj blisko daty, kolumny 1-3 w bloku)
        h_col = None
        for test_c in range(max(1, c_idx-10), c_idx):
            v = str(ws.cell(row=6, column=test_c).value).replace('.0','')
            if v.isdigit() and 7 <= int(v) <= 20:
                h_col = test_c
                break
        
        if not h_col: continue

        for row in range(5, ws.max_row + 1):
            try:
                cell = ws.cell(row=row, column=c_idx)
                val = str(cell.value).strip()
                if not val or val.lower() == "nan" or GRUPA in val.upper(): continue
                
                h = int(float(str(ws.cell(row=row, column=h_col).value)))
                m = int(float(str(ws.cell(row=row, column=h_col+1).value)))
                
                subject = " ".join(val.split())
                if is_red(cell): subject = f"🔴 [ZDALNIE] {subject}"
                
                all_slots.append({'start': current_date.replace(hour=h, minute=m), 'title': subject, 'col': c_idx})
            except: continue

    # 2. Łączenie slotów (45 min -> długie bloki)
    all_slots.sort(key=lambda x: (x['col'], x['start']))
    merged = []
    if all_slots:
        curr = all_slots[0].copy()
        curr['end'] = curr['start'] + timedelta(minutes=45)
        for nxt in all_slots[1:]:
            if nxt['title'] == curr['title'] and nxt['col'] == curr['col'] and nxt['start'] == curr['end']:
                curr['end'] = nxt['start'] + timedelta(minutes=45)
            else:
                merged.append(curr)
                curr = nxt.copy()
                curr['end'] = curr['start'] + timedelta(minutes=45)
        merged.append(curr)

    # 3. Generowanie ICS
    ics = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//PlanBot//MECHII", "X-WR-CALNAME:Plan MECH II", "METHOD:PUBLISH"]
    for e in merged:
        uid = hashlib.md5(f"{e['title']}{e['start']}{e['col']}".encode()).hexdigest()
        ics.extend(["BEGIN:VEVENT", f"UID:{uid}@studia.pl", f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}",
                    f"DTSTART:{e['start'].strftime('%Y%m%dT%H%M%S')}", f"DTEND:{e['end'].strftime('%Y%m%dT%H%M%S')}",
                    f"SUMMARY:{e['title']}", "END:VEVENT"])
    ics.append("END:VCALENDAR")
    
    with open(ICS_FILE, "w", encoding="utf-8") as f:
        f.write("\r\n".join(ics))

    # 4. Telegram
    merged.sort(key=lambda x: x['start'])
    report = "🚨 *PEŁNY PLAN MECH II*\n"
    last_d = ""
    for e in merged:
        d_str = e['start'].strftime("%d.%m")
        if d_str != last_d:
            report += f"\n📅 *{d_str}*\n"
            last_d = d_str
        report += f"  • {e['start'].strftime('%H:%M')}-{e['end'].strftime('%H:%M')}: {e['title']}\n"
    
    send_msg(report if merged else "⚠️ Nie znaleziono zajęć mimo znalezienia kolumn.")

if __name__ == "__main__":
    main()

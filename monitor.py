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
    if not cell.font or not cell.font.color: return False
    c = cell.font.color
    return str(c.rgb) in ["FFFF0000", "FF0000"] or c.indexed == 10

def main():
    print("🌐 Pobieranie strony...")
    r = requests.get(URL_STRONY)
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(r.text, 'html.parser')
    plan_url = next((a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx')), None)
    
    if not plan_url:
        print("❌ Nie znaleziono linku!")
        return

    print(f"📂 Pobieranie Excela: {plan_url}")
    resp = requests.get(plan_url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    
    wb = openpyxl.load_workbook("temp.xlsx", data_only=True)
    ws = wb.active
    print(f"✅ Arkusz otwarty. Max rows: {ws.max_row}")

    # 1. Szukanie dat i kolumn grupy
    all_blocks = []
    mech_cols_global = []

    for c in range(1, ws.max_column + 1):
        # Szukanie dat (wiersze 1-3)
        for r in range(1, 4):
            val = ws.cell(row=r, column=c).value
            d = parse_date(val)
            if d: all_blocks.append({"date": d, "col": c})
        
        # Szukanie kolumn grupy (wiersze 3-4)
        for r in range(3, 5):
            val = str(ws.cell(row=r, column=c).value).upper()
            if GRUPA in val: mech_cols_global.append(c)

    print(f"📅 Znaleziono dni: {len(all_blocks)}, Kolumn grupy: {len(mech_cols_global)}")

    all_slots = []
    for i, block in enumerate(all_blocks):
        c_start = block["col"]
        c_end = all_blocks[i+1]["col"] if i+1 < len(all_blocks) else ws.max_column + 1
        target_cols = [c for c in mech_cols_global if c_start <= c < c_end]
        
        # Szukanie godziny (kolumna z liczbami blisko daty)
        h_col = None
        for test_c in range(max(1, c_start-1), c_start+2):
            val = str(ws.cell(row=6, column=test_c).value).replace('.0','')
            if val.isdigit() and 7 <= int(val) <= 20:
                h_col = test_c
                break
        
        if not h_col: continue

        for row in range(5, ws.max_row + 1):
            try:
                h = int(float(str(ws.cell(row=row, column=h_col).value)))
                m = int(float(str(ws.cell(row=row, column=h_col+1).value)))
                time_dt = block["date"].replace(hour=h, minute=m)
                
                for t_c in target_cols:
                    cell = ws.cell(row=row, column=t_c)
                    val = " ".join(str(cell.value).split())
                    if not val or val.lower() == "nan" or GRUPA in val.upper(): continue
                    
                    if is_red(cell): val = f"🔴 [ZDALNIE] {val}"
                    all_slots.append({'start': time_dt, 'title': val, 'col': t_c})
            except: continue

    # 2. Łączenie slotów (45min -> długie bloki)
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

    print(f"✅ Wyłuskano {len(merged)} pełnych bloków zajęć.")

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
    report = "🚨 *PLAN MECH II (Zdalne + Godziny)*\n"
    last_d = ""
    for e in merged:
        d_str = e['start'].strftime("%d.%m")
        if d_str != last_d:
            report += f"\n📅 *{d_str}*\n"
            last_d = d_str
        report += f"  • {e['start'].strftime('%H:%M')}-{e['end'].strftime('%H:%M')}: {e['title']}\n"
    
    send_msg(report)

if __name__ == "__main__":
    main()

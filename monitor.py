import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import json
from datetime import datetime

# --- KONFIGURACJA ---
URL_STRONY = "https://uczelniaoswiecim.edu.pl/instytuty/new-instytut-nauk-inzynieryjno-technicznych/harmonogramy/"
TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")
STATE_FILE = "plan_state.json"
GRUPA = "MECH II"

def send_msg(text):
    if not TOKEN or not CHAT_ID: return
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    # Dzielimy na mniejsze partie, jeśli tekst jest za długi
    if len(text) > 4000: text = text[:4000] + "..."
    requests.post(url, data={"chat_id": CHAT_ID, "text": text, "parse_mode": "Markdown"})

def get_plan_url():
    r = requests.get(URL_STRONY)
    soup = BeautifulSoup(r.text, 'html.parser')
    for a in soup.find_all('a', href=True):
        if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx'):
            return a['href']
    return None

def parse_excel(url):
    resp = requests.get(url)
    with open("temp.xlsx", "wb") as f:
        f.write(resp.content)
    
    # Czytamy Excela
    df = pd.read_excel("temp.xlsx", header=None)
    
    # Szukamy kolumn MECH II (zazwyczaj wiersz 3, czyli index 2)
    mech_cols = []
    for i, val in enumerate(df.iloc[2]):
        if str(val).strip() == GRUPA:
            mech_cols.append(i)
    
    events = []
    # Przeszukujemy wiersze od 4 w górę
    for col in mech_cols:
        current_date = "Nieokreślona"
        for row in range(0, len(df)):
            cell_val = str(df.iloc[row, col]).strip()
            
            # Próba wyłapania daty (zazwyczaj wiersz 0 lub 1 nad grupą)
            potential_date = str(df.iloc[0, col-2] if col > 2 else "") 
            
            if cell_val != "nan" and cell_val != "" and cell_val != GRUPA:
                # Tworzymy unikalny klucz wydarzenia: Wiersz_Kolumna_Zawartość
                # To najbezpieczniejszy sposób na wykrycie przesunięć w tabeli
                events.append(f"{cell_val}")
                
    return sorted(list(set(events)))

def main():
    print("🚀 Analiza szczegółowa planu MECH II...")
    url = get_plan_url()
    if not url: return

    current_events = parse_excel(url)
    
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r", encoding='utf-8') as f:
            old_events = json.load(f)
    else:
        old_events = []

    # Porównanie
    added = [e for e in current_events if e not in old_events]
    removed = [e for e in old_events if e not in current_events]

    if added or removed:
        report = "🔔 *ZMIANY W HARMONOGRAMIE MECH II!*\n\n"
        
        if added:
            report += "✅ *NOWE / ZMIENIONE:*\n"
            for a in added[:10]: # Limit 10, żeby nie spamować
                report += f"• {a}\n"
        
        if removed:
            report += "\n❌ *USUNIĘTE / STARE:*\n"
            for r in removed[:10]:
                report += f"• {r}\n"
        
        report += f"\n🔗 [Pobierz pełny Excel]({url})"
        send_msg(report)
        
        # Zapisujemy nowy stan
        with open(STATE_FILE, "w", encoding='utf-8') as f:
            json.dump(current_events, f, ensure_ascii=False)
    else:
        print("Brak istotnych zmian w treści zajęć.")

if __name__ == "__main__":
    main()

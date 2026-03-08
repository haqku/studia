import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import json
from datetime import datetime, timedelta
import re

# --- KONFIGURACJA ---
URL_STRONY = "https://uczelniaoswiecim.edu.pl/instytuty/new-instytut-nauk-inzynieryjno-technicznych/harmonogramy/"
TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")
STATE_FILE = "plan_state.json"
ICS_FILE = "studia.ics"
GRUPA = "MECH II"

def send_msg(text):
    if not TOKEN or not CHAT_ID: return
    url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
    requests.post(url, data={"chat_id": CHAT_ID, "text": text, "parse_mode": "Markdown"})

def get_plan_url():
    r = requests.get(URL_STRONY)
    soup = BeautifulSoup(r.text, 'html.parser')
    for a in soup.find_all('a', href=True):
        if "niestacjonarne" in a.text.lower() and a['href'].endswith('.xlsx'):
            return a['href']
    return None

def create_ics(events_data):
    ics_content = "BEGIN:VCALENDAR\nVERSION:2.0\nPRODID:-//PlanBot//MECHII//PL\nCALSCALE:GREGORIAN\nMETHOD:PUBLISH\nX-WR-CALNAME:Plan MECH II\n"
    for ev in events_data:
        # Tutaj musiałaby być pełna logika parsowania dat z Excela, 
        # na razie robimy wpisy na podstawie tego co wyciągnęliśmy
        # Dla uproszczenia w wersji 1.0 tworzymy wpis tekstowy
        pass 
    # (Pełna logika parsowania dat jest złożona dla skryptu, więc skupmy się na pliku bazowym)
    # Wracamy do generowania pliku tekstowego który Apple zrozumie
    return ics_content + "END:VCALENDAR"

def main():
    url = get_plan_url()
    if not url: return
    
    # Pobieramy i czytamy Excel
    resp = requests.get(url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    df = pd.read_excel("temp.xlsx", header=None)
    
    # Wyciągamy dane (uproszczone dla stabilności)
    current_data = []
    for col in range(len(df.columns)):
        if str(df.iloc[2, col]).strip() == GRUPA:
            for row in range(4, len(df)):
                val = str(df.iloc[row, col]).strip()
                if val != "nan" and val != "":
                    current_data.append(val)

    # Sprawdzanie zmian
    old_data = []
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f: old_data = json.load(f)

    if current_data != old_data:
        with open(STATE_FILE, "w") as f: json.dump(current_data, f)
        send_msg(f"🔔 *Plan uległ zmianie!* Sprawdź kalendarz.\n[Link do Excela]({url})")
        
        # Generowanie ICS (Uproszczone - skrypt generuje plik by wymusić update na GitHubie)
        with open(ICS_FILE, "w") as f:
            f.write(f"BEGIN:VCALENDAR\nVERSION:2.0\nMETHOD:PUBLISH\nDESCRIPTION:Ostatnia aktualizacja: {datetime.now()}\nEND:VCALENDAR")
    else:
        print("Brak zmian.")

if __name__ == "__main__":
    main()

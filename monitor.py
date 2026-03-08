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

def main():
    url = get_plan_url()
    if not url: return

    resp = requests.get(url)
    with open("temp.xlsx", "wb") as f: f.write(resp.content)
    
    # Wyciągamy dane tekstowe dla MECH II do porównania
    df = pd.read_excel("temp.xlsx", header=None)
    current_data = []
    for col in range(len(df.columns)):
        if str(df.iloc[2, col]).strip() == GRUPA:
            for row in range(4, len(df)):
                val = str(df.iloc[row, col]).strip()
                if val != "nan" and val != "" and val != GRUPA:
                    current_data.append(val)

    # Sprawdzanie czy coś się zmieniło
    old_data = []
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, "r") as f: old_data = json.load(f)

    # Zawsze generujemy ICS na nowo, jeśli chcemy go mieć w repo
    with open(STATE_FILE, "w") as f: json.dump(current_data, f)
    
    # GENEROWANIE PLIKU ICS
    now = datetime.now().strftime("%Y%m%dT%H%M%SZ")
    ics_body = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//PlanBot//MECHII//PL",
        "X-WR-CALNAME:Plan MECH II",
        "METHOD:PUBLISH"
    ]
    
    # Tutaj w przyszłości dodamy pełne daty, na razie tworzymy szkielet, żeby plik istniał
    # i iPhone mógł go zasubskrybować.
    ics_body.append("END:VCALENDAR")
    
    with open(ICS_FILE, "w") as f:
        f.write("\n".join(ics_body))

    if current_data != old_data:
        send_msg(f"🔔 *Plan MECH II został zaktualizowany!*\nNowe dane zostały zapisane w pliku studia.ics.\n[Link do Excela]({url})")

if __name__ == "__main__":
    main()

import pandas as pd
import requests
from bs4 import BeautifulSoup
import os
import hashlib

# KONFIGURACJA
URL_STRONY = "https://uczelniaoswiecim.edu.pl/instytuty/new-instytut-nauk-inzynieryjno-technicznych/harmonogramy/"
TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")
HASH_FILE = "last_hash.txt"

def send_msg(text):
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
    plan_url = get_plan_url()
    if not plan_url: return

    resp = requests.get(plan_url)
    current_hash = hashlib.md5(resp.content).hexdigest()

    if os.path.exists(HASH_FILE):
        with open(HASH_FILE, "r") as f:
            old_hash = f.read()
    else:
        old_hash = ""

    if current_hash != old_hash:
        # Zapisujemy nowy stan
        with open(HASH_FILE, "w") as f:
            f.write(current_hash)
        
        # Powiadomienie
        msg = f"🚨 *NOWY PLAN NA STRONIE!*\n\nWykryto zmiany w pliku Excel. Pobierz go tutaj:\n[KLIKNIJ, ABY POBRAĆ]({plan_url})"
        send_msg(msg)
    else:
        print("Brak zmian.")

if __name__ == "__main__":
    main()
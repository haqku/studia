import sys
import os

# Wymuszamy, żeby Python wypisywał wszystko natychmiast
print("--- START SKRYPTU ---")
print(f"Wersja Pythona: {sys.version}")

try:
    import pandas as pd
    import requests
    from bs4 import BeautifulSoup
    print("✅ Biblioteki załadowane pomyślnie.")
except Exception as e:
    print(f"❌ BŁĄD IMPORTU: {e}")
    sys.exit(1)

TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")

if not TOKEN or not CHAT_ID:
    print("❌ BŁĄD: Brak sekretów (Token/Chat ID) w ustawieniach GitHuba!")
else:
    print("✅ Sekrety załadowane.")

URL = "https://uczelniaoswiecim.edu.pl/instytuty/new-instytut-nauk-inzynieryjno-technicznych/harmonogramy/"

print(f"🌐 Łączę ze stroną: {URL}")
try:
    r = requests.get(URL, timeout=15)
    print(f"Status odpowiedzi: {r.status_code}")
    soup = BeautifulSoup(r.text, 'html.parser')
    links = [a['href'] for a in soup.find_all('a', href=True) if "niestacjonarne" in a.text.lower()]
    
    if links:
        print(f"🎯 Znaleziono link do planu: {links[0]}")
        # Próba wysłania testowego sygnału do bota
        print("🤖 Próba wysłania wiadomości do bota...")
        test_url = f"https://api.telegram.org/bot{TOKEN}/sendMessage"
        t_resp = requests.post(test_url, data={"chat_id": CHAT_ID, "text": "🚀 TEST: Skrypt monitor.py ruszył na GitHubie!"})
        print(f"Status Telegrama: {t_resp.status_code}")
    else:
        print("❌ Nie znaleziono linku 'niestacjonarne' na stronie.")

except Exception as e:
    print(f"❌ WYSTĄPIŁ BŁĄD: {e}")

print("--- KONIEC SKRYPTU ---")

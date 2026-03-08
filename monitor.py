name: Monitor Harmonogramu
on:
  schedule:
    - cron: '0 */4 * * *' # Sprawdzaj co 4 godziny
  workflow_dispatch:      # Pozwala odpalić ręcznie

jobs:
  run-monitor:
    runs-on: ubuntu-latest
    permissions:
      contents: write     # Dodatkowe zabezpieczenie uprawnień
    steps:
      - uses: actions/checkout@v3

      - name: Setup Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.9'

      - name: Install dependencies
        run: pip install pandas openpyxl requests beautifulsoup4

      - name: Run script
        env:
          TELEGRAM_TOKEN: ${{ secrets.TELEGRAM_TOKEN }}
          CHAT_ID: ${{ secrets.CHAT_ID }}
        run: python monitor.py

      - name: Commit and Push changes
        run: |
          git config --global user.name 'PlanBot'
          git config --global user.email 'bot@github.com'
          git add plan_state.json studia.ics last_hash.txt || true
          git commit -m "Automatyczna aktualizacja planu i kalendarza" || exit 0
          git push

# Cinema Program Checker

Webová aplikácia na porovnanie týždenného Cinema City programu s interným XLSX rozpisom.

## Čo kontroluje
- kino podľa hárku: Eurovea / Aupark / Polus
- celý programový týždeň štvrtok–streda
- `only` a `w/o` pravidlá v časových bunkách
- film a čas
- atribúty/formáty
- SUB/DUB vrátane jazyka

Sála sa číta z Excelu a zobrazuje pri chybe. Verejný Cinema City listing ju nemusí poskytovať, preto ju aplikácia nevyhlasuje za chybnú bez webového údaja.

## Lokálne spustenie
1. `pip install -r requirements.txt`
2. `playwright install chromium`
3. `python app.py`
4. Otvor `http://localhost:5000`

## Render
Repozitár obsahuje `Dockerfile` a `render.yaml`. Po nahratí do GitHubu vytvor na Renderi nový Blueprint/Web Service z repozitára.

Poznámka: Cinema City je dynamický web a môže meniť HTML/ochranu. Scraper preto môže v budúcnosti vyžadovať úpravu selektorov/parsingovej logiky. Aplikácia pri nulových webových dátach kontrolu zastaví namiesto vytvorenia falošných chýb.

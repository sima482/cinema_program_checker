# Cinema Program Checker

Webová aplikácia na porovnanie interného XLSX programu s verejným programom Cinema City Slovensko.

## Kontrola
- Eurovea / Aupark / Polus
- celý týždeň štvrtok–streda
- film + čas
- atribúty/formáty
- SUB/DUB vrátane jazyka
- pravidlá `only` a `w/o`

Aplikácia má aj náhľad „Ukázať, čo čítam z Excelu“, aby sa dala najprv overiť interpretácia tabuľky bez kontaktovania Cinema City.

Ak Cinema City nevráti program pre deň, v ktorom Excel očakáva predstavenia, kontrola sa zastaví. Tak nevznikne veľký zoznam falošných chýb.

## Render
Projekt používa Docker + Playwright. `render.yaml` obsahuje health check `/health`.

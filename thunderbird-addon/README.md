# MailAI Thunderbird add-on (experimental)

Tento adresar obsahuje prvni verzi Thunderbird add-onu, ktery prenasi hlavni funkce z `local_app.py` do Thunderbirdu.

`local_app.py` zustava beze zmen.

## Co add-on umi

- nacist neprectene emaily za poslednich N dni
- poslat souhrn do OpenAI-compatible endpointu
- rozdelit emaily do bucketu: urgentni, stredne_dulezite, pocka, k_preposlani, ignorovat
- aplikovat Thunderbird tagy podle bucketu
- volitelne pridat deadline tag
- provest doporucene `mark_read_ids`

## Co je jinak oproti Outlook/Graph verzi

- nepouziva Microsoft Graph
- nepouziva Outlook masterCategories ani kalendarove udalosti
- pracuje s Thunderbird tags a local storage

## Instalace pro vyvoj

1. Otevri Thunderbird.
2. Otevri `Add-ons and Themes`.
3. Vyber `Debug Add-ons`.
4. Klikni `Load Temporary Add-on...`.
5. Vyber soubor `thunderbird-addon/manifest.json`.

## Poznamky

- API key je ulozen v `browser.storage.local` v ramci add-onu.
- Add-on je experimentalni a muze vyzadovat drobne upravy podle konkretni Thunderbird verze.
- Pokud vase verze nepodporuje nektere `browser.messages.*` metody, je potreba API adaptace.

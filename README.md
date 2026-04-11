# MailAI

MailAI je Outlook add-in pro asistované třídění e-mailů pomocí LLM (primárně e-infra OpenAI-compatible API).

Aplikace obsahuje:
- frontend add-in v Outlook taskpane ([taskpane.html](taskpane.html))
- Python backend ve Flasku ([backend.py](backend.py))
- Office add-in manifest ([manifest.xml](manifest.xml))

## Funkce

- Analýza otevřeného e-mailu (priorita, kategorie, shrnutí, návrh odpovědi)
- Inbox souhrn nepřečtených e-mailů za posledních N dní (default 10)
- Vlastní pravidla přes custom prompt
- Preferovaní odesílatelé (priority boost v promptu)
- Dynamické načítání dostupných modelů z provideru
- Uložení nastavení do Outlook RoamingSettings
- Vložení návrhu odpovědi do reply formuláře (pokud to klient Outlooku podporuje)

## Architektura

1. Outlook add-in načte obsah taskpane stránky.
2. Uživatel v Nastavení zadá:
- Backend URL
- volitelný Backend API key (pro ochranu backend endpointů)
- LLM API key (povinný pro volání modelu)
- LLM Base URL
- model
3. Frontend posílá data na backend endpointy.
4. Backend volá OpenAI-compatible API s klíčem dodaným uživatelem v requestu.

Poznámka: LLM API key není uložen na backendu jako server env proměnná.

## Struktura projektu

- [backend.py](backend.py): Flask API, servování taskpane, volání LLM provideru
- [taskpane.html](taskpane.html): Outlook add-in UI + Office.js logika
- [manifest.xml](manifest.xml): Office Add-in manifest
- [requirements.txt](requirements.txt): Python závislosti
- [render.yaml](render.yaml): Render deploy konfigurace

## API endpointy backendu

- GET [/](/)
	- vrací taskpane HTML

- GET [/taskpane.html](/taskpane.html)
	- vrací taskpane HTML

- GET [/health](/health)
	- health check endpoint
	- response: `{ "status": "ok" }`

- GET [/assets/<filename>](/assets/icon-64.png)
	- statické assety
	- pokud nejsou dostupné ikony, backend vrací transparentní PNG fallback pro `icon-64.png` a `icon-128.png`

- POST [/models](/models)
	- načte dostupné modely z aktuálně nastaveného provideru
	- request body očekává minimálně:
		- `llmApiKey`
		- volitelně `llmBaseUrl`
	- response:
		- `models`: pole názvů modelů
		- `defaultModel`: backend default model

- POST [/analyze](/analyze)
	- analyzuje jeden e-mail
	- request body (hlavní pole):
		- `subject`, `from`, `date`, `body`
		- `model`
		- `llmApiKey` (povinné)
		- `llmBaseUrl` (volitelné)
		- `customPrompt`, `prioritySenders`

- POST [/analyze-inbox](/analyze-inbox)
	- načte nepřečtené e-maily z Inboxu přes Microsoft Graph za posledních N dní a vrátí souhrn
	- request body (hlavní pole):
		- `token` (Graph access token)
		- `days` (default 10)
		- `top`
		- `model`
		- `llmApiKey` (povinné)
		- `llmBaseUrl` (volitelné)
		- `customPrompt`, `prioritySenders`
	- response obsahuje kategorie:
		- `urgentni`
		- `stredne_dulezite`
		- `pocka`
		- `k_preposlani`
		- `ignorovat`

## Autentizace

### Backend autentizace (volitelná)

Pokud je nastavena env proměnná `API_SECRET`, backend vyžaduje hlavičku:

`Authorization: Bearer <API_SECRET>`

Tohle je oddělené od LLM API klíče.

### LLM autentizace

LLM API key posílá frontend v JSON payloadu jako `llmApiKey`.

## Konfigurace

### Backend env proměnné

- `EINFRA_BASE_URL` (default: `https://llm.ai.e-infra.cz/v1/`)
- `EINFRA_MODEL` (default model použitý pokud klient nepošle `model`)
- `API_SECRET` (volitelné, ochrana backend endpointů)

### Frontend nastavení (Outlook RoamingSettings)

- `apiUrl`
- `backendApiKey`
- `apiProvider`
- `llmApiKey`
- `llmBaseUrl`
- `model`
- `graphToken`
- `inboxDays`
- `inboxTop`
- `customPrompt`
- `prioritySenders`
- `autoload`

## Lokální spuštění

Požadavky:
- Python 3.10+

Instalace:

```bash
pip install -r requirements.txt
```

Spuštění:

```bash
python backend.py
```

Backend běží standardně na `http://localhost:5000`.

## Lokální aplikace bez Outlook add-inu

Pokud je ve firmě zakázané nahrávání Outlook add-inů, použij lokální aplikaci [local_app.py](local_app.py).

Co umí:
- načíst nepřečtené e-maily za posledních N dní
- udělat AI souhrn v kategoriích urgentní / středně důležité / počká / k přeposlání / ignorovat
- nabídnout doporučené hromadné akce (označit jako přečtené, smazat)

Spuštění:

```bash
pip install -r requirements.txt
streamlit run local_app.py
```

V aplikaci vyplň:
1. LLM API key
2. LLM Base URL
3. Model
4. Graph Access Token
5. Počet dní zpět (typicky 10)

Poznámka:
- Pro samotné čtení inboxu stačí obvykle `Mail.Read`.
- Pro hromadné akce (označit přečtené / smazat) je typicky potřeba `Mail.ReadWrite`.

## Deploy na Render

Repo je připraven pro Render Web Service přes [render.yaml](render.yaml).

Start command:

```bash
gunicorn backend:app --bind 0.0.0.0:$PORT
```

Doporučení:
- nastav Health Check Path na `/health`
- pokud chceš chránit backend, nastav `API_SECRET`

## Outlook manifest

V [manifest.xml](manifest.xml) musí být URL na produkční host.

Aktuálně je nastaveno na:
- `https://mailai-kr2k.onrender.com/taskpane.html`
- ikony: `https://mailai-kr2k.onrender.com/assets/icon-64.png` a `.../icon-128.png`

Po změně domény uprav manifest a znovu proveď sideload add-inu.

## Jak přidat add-in do Outlooku

### Outlook on the web

1. Otevři Outlook on the web.
2. Klikni na `Aplikace` nebo `Získat doplňky`.
3. Otevři správu doplňků (`My add-ins` / `Moje doplňky`).
4. Zvol přidání vlastního doplňku ze souboru.
5. Nahraj [manifest.xml](manifest.xml).
6. Potvrď instalaci.

### Outlook desktop

1. Otevři Outlook.
2. Přejdi do nabídky pro doplňky (`Get Add-ins` / `Získat doplňky`).
3. Otevři `My add-ins` / `Moje doplňky`.
4. Vyber přidání vlastního doplňku ze souboru.
5. Nahraj [manifest.xml](manifest.xml).

### Kdy je potřeba add-in nahrát znovu

Manifest znovu nahraj, pokud změníš:
- URL v [manifest.xml](manifest.xml)
- název, ikony nebo jiné manifest metadata

Pouhý redeploy backendu nebo změna obsahu [taskpane.html](taskpane.html) obvykle nevyžaduje nový upload manifestu, pokud zůstává stejná URL.

## Práce s modely

V Nastavení je tlačítko Načíst modely:

1. Frontend zavolá `/models` s uživatelským LLM klíčem.
2. Backend načte aktuální modely z provideru.
3. UI doplní model suggestions dynamicky podle reálné dostupnosti.

## Microsoft Graph token pro Inbox souhrn

Inbox souhrn potřebuje Graph Access Token s oprávněním číst poštu.

### Nejrychlejší způsob (Graph Explorer)

1. Otevři https://developer.microsoft.com/graph/graph-explorer
2. Přihlas se stejným Microsoft 365 účtem, který používáš v Outlooku.
3. Klikni vpravo nahoře na profilový avatar.
4. Zvol `Consent to permissions`.
5. Vyhledej `Mail.Read`.
6. U `Mail.Read` potvrď `Consent`.
7. Zavři dialog a v hlavní query oblasti otevři záložku `Access token`.
8. Zkopíruj hodnotu access tokenu.

Poznámka:
- Token je časově omezený. Po expiraci je potřeba získat nový.
- Kopíruj pouze token samotný, ne celý řetězec `Bearer ...`.
- Záložka `Access token` je v Graph Exploreru v request/query části vedle dalších záložek jako request body nebo modify permissions.

### Kde přesně token najdu

Po udělení oprávnění je token dostupný přímo v Graph Exploreru:

1. Zůstaň přihlášený v Graph Exploreru.
2. V prostřední části stránky, kde se skládá request, najdi záložku `Access token`.
3. Otevři ji.
4. Zobrazí se OAuth access token pro aktuálně přihlášený účet.
5. Zkopíruj celý token a vlož ho do pluginu do pole `Graph Access Token`.

Pokud záložku `Access token` nevidíš:
- spusť nejdřív libovolný Graph dotaz, například `GET /me/messages?$top=1`
- zkontroluj, že jsi přihlášený
- zkus obnovit stránku po consentu

### Alternativa (vlastní Azure App Registration)

Použij pokud nechceš Graph Explorer nebo potřebuješ produkční flow.

1. V Azure Portal vytvoř App Registration.
2. Přidej delegated permission `Mail.Read` pro Microsoft Graph.
3. Uděl admin/user consent.
4. Získej access token přes OAuth 2.0 (authorization code flow).
5. Token použij stejně jako v pluginu.

## Postup vyplnění tokenu v pluginu

1. Otevři add-in a přejdi na záložku Inbox souhrn.
2. Do pole Graph Access Token vlož získaný token (bez prefixu `Bearer`).
3. Nastav Počet dní zpět (např. 10).
4. Nastav Max. počet e-mailů (např. 200).
5. Klikni Analyzovat nepřečtené e-maily.

Tip:
- Když backend vrátí `InvalidAuthenticationToken`, token je neplatný nebo expirovaný.
- Když vrátí `AccessDenied`, chybí scope `Mail.Read` nebo consent.

## Troubleshooting

### Deploy na Render trvá dlouho

- ověř Health Check Path (`/health`)
- zkontroluj log, že Gunicorn opravdu naslouchá na `$PORT`
- ověř, že build neskončil na instalaci závislostí

### Chyba Missing LLM API key

- v add-inu vyplň pole LLM API key
- ulož nastavení

### /taskpane.html vrací 404

- ověř, že běží aktuální verze backendu obsahující route pro `/taskpane.html`
- udělej redeploy

### Modely se nenačtou

- ověř `llmBaseUrl` a `llmApiKey`
- zkontroluj, že provider podporuje endpoint models list

## Bezpečnostní poznámky

- LLM API key je citlivý údaj, zvaž použití jen pro session.
- Pokud používáš `API_SECRET`, neposílej backend endpointy bez autorizace veřejně.
- Před produkčním nasazením zvaž audit logování, rate limiting a CORS policy.

## Ochrana backendu přes API_SECRET

### K čemu je ochrana backendu

`API_SECRET` chrání backend endpointy před neautorizovaným použitím z internetu.

Přínosy:
- omezení zneužití veřejného endpointu
- menší riziko spam requestů a přetížení
- kontrola, kdo může volat `models`, `analyze`, `analyze-inbox`

Kdy použít:
- produkce: doporučeno zapnout
- lokální vývoj / krátké demo: může být dočasně vypnuto

### Kam přesně nahrát API_SECRET (Render)

1. Otevři Render Dashboard.
2. Vyber službu `mailai-kr2k`.
3. Jdi do Environment -> Environment Variables.
4. Přidej proměnnou:
	- Key: `API_SECRET`
	- Value: silný náhodný tajný řetězec
5. Ulož změnu a spusť redeploy.

Poznámka:
- `API_SECRET` nikdy neukládej do repozitáře ani přímo do kódu.

### Co vyplnit v pluginu

1. Otevři záložku Nastavení v add-inu.
2. Do pole Backend API key vlož stejnou hodnotu jako `API_SECRET` na Renderu.
3. Ulož nastavení.

Plugin pak automaticky přidává hlavičku:
- `Authorization: Bearer <API_SECRET>`

### Rotace tajného klíče

Při podezření na únik:
1. změň `API_SECRET` v Renderu
2. redeploy
3. aktualizuj Backend API key v pluginu

## Licence

Interní projekt / bez explicitní open-source licence.

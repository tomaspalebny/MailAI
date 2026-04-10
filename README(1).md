# 📬 Email Triage AI — Outlook Add-in

AI asistent pro třídění, sumarizaci a návrhy odpovědí na e-maily přímo v Outlooku.

## Architektura

```
┌─────────────────┐     ┌──────────────────┐     ┌─────────────┐
│  Outlook Add-in │────▶│  Python Backend   │────▶│  OpenAI API │
│  (taskpane.html)│◀────│  (Render.com)     │◀────│  (GPT-4o)   │
│  Office.js      │     │  Flask + CORS     │     │             │
└─────────────────┘     └──────────────────┘     └─────────────┘
                              │
                              ▼
                        ┌─────────────┐
                        │ Graph API   │  (pro inbox sumář)
                        └─────────────┘
```

## Rychlý start

### 1. Deploy backend na Render.com

1. Vytvořte nový Web Service na render.com
2. Nahrajte `backend.py`, `requirements.txt` a `render.yaml`
3. Nastavte env proměnné:
   - `EINFRA_API_KEY` = váš API klíč pro llm.ai.e-infra.cz
   - `OPENAI_API_KEY` = fallback varianta (pokud nepoužijete `EINFRA_API_KEY`)
   - `EINFRA_BASE_URL` = `https://llm.ai.e-infra.cz/v1/`
   - `EINFRA_MODEL` = `gpt-4o-mini` (volitelné)
   - `API_SECRET` = libovolný tajný klíč (volitelný)
4. Start command: `gunicorn backend:app --bind 0.0.0.0:$PORT`
5. Poznamenejte si URL (např. https://email-triage-backend.onrender.com)

### 2. Aktualizujte manifest

V `manifest.xml` nahraďte všechny výskyty `YOUR-SERVER.onrender.com`
vaší skutečnou URL z Render.com.

### 3. Hostujte taskpane.html

Soubor `taskpane.html` musí být dostupný na HTTPS. Možnosti:
- **Render.com Static Site** — nejsnazší, nahrajte taskpane.html
- **GitHub Pages** — zdarma HTTPS hosting
- **Váš server** — jakýkoli HTTPS endpoint

### 4. Sideload do Outlooku

#### Outlook on the web (nejsnazší pro testování):
1. Otevřete outlook.office.com
2. Klikněte na ⚙️ → Spravovat doplňky → Moje doplňky
3. "Přidat vlastní doplněk" → "Přidat ze souboru"
4. Nahrajte `manifest.xml`

#### Outlook desktop (Windows):
1. Soubor → Spravovat doplňky
2. Vlastní doplňky → Přidat ze souboru
3. Nahrajte `manifest.xml`

#### Admin deployment (pro celou organizaci MUNI):
1. Microsoft 365 Admin Center → Nastavení → Integrované aplikace
2. Nahrát manifest → nasadit pro vybrané uživatele/skupiny

### 5. Použití

1. Otevřete e-mail v Outlooku
2. Na ribbonu klikněte "Analyzovat e-mail" (📬)
3. V panelu vpravo se zobrazí priorita, souhrn a návrh odpovědi
4. V záložce Pravidla upravte custom prompt a preferované odesílatele

## Nastavení v add-inu

V záložce ⚙️ nastavte:
- **Backend URL** — adresa vašeho backendu na Render.com
- **API klíč** — pokud jste nastavili API_SECRET
- **AI Model** — GPT-4o (nejlepší), GPT-4o Mini (levnější)
- **Jazyk odpovědí** — čeština, angličtina, nebo auto

## Náklady

- **Render.com**: Free tier (750h/měsíc) stačí pro testování
- **OpenAI API**: ~$0.01 za jeden e-mail, ~$0.10 za inbox 50 e-mailů
- **Outlook Add-in**: zdarma (sideloading nebo admin deployment)

## Bezpečnost

⚠️ E-maily se posílají na váš backend a odtud na OpenAI API.
Pro citlivé prostředí zvažte:
- **Azure OpenAI** místo veřejného OpenAI (data zůstanou v EU)
- **Ollama** na lokálním serveru (změňte base_url v backendu)

## Rozšíření

- [ ] Automatické spuštění při otevření e-mailu (event-based activation)
- [ ] Tlačítko "Odpovědět" — vloží návrh přímo do reply okna
- [ ] Denní digest e-mailem přes scheduled task
- [ ] Učení z uživatelových oprav priorit (fine-tuning)

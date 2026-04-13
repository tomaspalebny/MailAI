# backend_v2

Alternativni backend verze postavena podle funkcionalit z local_app.

Dulezite:
- local_app zustava beze zmen
- puvodni backend.py zustava beze zmen
- tato verze je izolovana v novem adresari

## Co backend_v2 pridava

- analyza inboxu v rezimech:
  - Neprectene
  - Bez odpovedi (Inbox vs Sent)
- filtrace uz zpracovanych MailAI zprav:
  - Jen bez MailAI stitku + urgentni pripomenuti
  - Jen bez MailAI stitku
  - Vsechny (vcetne jiz oznacenych)
- obohaceni vysledku o metadata zpravy (receivedDateTime, webLink, categories)
- no-delete policy
- Graph diagnostika opravneni
- nacitani Outlook master categories
- aplikace klasifikace (prirazeni kategorii + optional deadline stitek + mark read)
- vytvareni kalendarove udalosti

## Endpointy

- GET /health
- POST /models
- POST /analyze
- POST /analyze-inbox
- POST /graph/diagnostics
- POST /graph/categories
- POST /apply-classification
- POST /calendar/create-event

## Spusteni

Pouziva stejne zavislosti jako root requirements.txt.

```bash
python backend_v2/backend.py
```

Server bezi na:

- http://localhost:5001

## Poznamka k frontendu

Puvodni taskpane muze dal volat /models, /analyze a /analyze-inbox.
Nove endpointy umoznuji dodelat funkcionality lokalni Streamlit verze i do web/add-in klienta.

# Outlook add-in pro e-infra.cz API

## Co je nové
- Custom prompt přímo v add-inu
- Seznam preferovaných odesílatelů
- Backend napojený na e-infra.cz API endpoint `https://llm.ai.e-infra.cz/v1/`

## Nastavení backendu
Proměnné prostředí:
- `EINFRA_API_KEY`
- `OPENAI_API_KEY` (kompatibilní fallback, pokud nepoužijete `EINFRA_API_KEY`)
- `EINFRA_BASE_URL=https://llm.ai.e-infra.cz/v1/`
- `EINFRA_MODEL=gpt-4o-mini`
- `API_SECRET` volitelně

## Poznámka k přístupu
API klíč pro e-infra získáte v prostředí chat.ai.e-infra.cz / llm.ai.e-infra.cz a používáte jej jako Bearer token.

## Acknowledgement
Tento balík je připraven pro napojení na e-infra.cz API a custom prompt v add-inu.

## Doplněno
- RoamingSettings pro uložení promptu a pravidel
- Automatické načtení nastavení při startu
- Šablony promptů pro univerzitu, IT a výuku

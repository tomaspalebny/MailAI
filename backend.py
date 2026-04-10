"""
Email Triage AI — Backend (Flask)
Deploy na Render.com jako Web Service.

Požadavky:
  pip install flask flask-cors openai requests gunicorn

Env proměnné na Render.com:
  OPENAI_API_KEY=sk-...
  API_SECRET=váš-tajný-klíč (volitelný)
"""

import os
import json
from flask import Flask, request, jsonify
from flask_cors import CORS
from openai import OpenAI
import requests as http_requests

app = Flask(__name__)
CORS(app)  # Outlook add-in posílá requesty z jiné domény

client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
API_SECRET = os.environ.get("API_SECRET", "")

SYSTEM_PROMPT = """Jsi e-mailový asistent na české univerzitě (Masarykova univerzita).
Analyzuj e-mail a vrať JSON objekt (ne pole, jeden objekt):
{
  "priority": "P0_urgent|P1_action_needed|P2_informational|P3_ignorable",
  "category": "administrativa|výuka|IT|výzkum|osobní|spam|newsletter",
  "summary": "1-2 věty česky, co e-mail chce a od koho",
  "needs_reply": true/false,
  "reply_deadline": "dnes|tento_tyden|zadny",
  "suggested_reply": "Návrh odpovědi v požadovaném jazyce pokud needs_reply=true, jinak null"
}

Pravidla pro prioritu:
- P0: deadline dnes/zítra, žádost od vedení/děkana, urgentní IT incident
- P1: vyžaduje akci tento týden — žádosti studentů, schvalování, úkoly
- P2: informativní — oznámení, newslettery, info bez nutnosti akce
- P3: spam, masové rozesílky, marketing, automatické notifikace

Návrh odpovědi:
- Formální český jazyk, pokud není řečeno jinak
- Stručný ale zdvořilý
- Obsahuje konkrétní reakci na obsah e-mailu"""

INBOX_SYSTEM_PROMPT = """Jsi e-mailový asistent na české univerzitě.
Dostaneš seznam e-mailů. Pro každý vrať JSON objekt.
Vrať JSON s klíčem "emails" obsahující pole objektů:
{
  "emails": [
    {
      "id": "ID e-mailu",
      "subject": "předmět",
      "priority": "P0_urgent|P1_action_needed|P2_informational|P3_ignorable",
      "category": "administrativa|výuka|IT|výzkum|osobní|spam|newsletter",
      "summary": "1 věta česky",
      "needs_reply": true/false,
      "reply_deadline": "dnes|tento_tyden|zadny",
      "suggested_reply": "Stručný návrh odpovědi nebo null"
    }
  ]
}
Řaď od nejvyšší priority (P0 nahoře). U P0 a P1 vždy navrhni odpověď."""


def check_auth():
    """Volitelná autorizace přes Bearer token."""
    if not API_SECRET:
        return True
    auth = request.headers.get("Authorization", "")
    return auth == f"Bearer {API_SECRET}"


@app.route("/analyze", methods=["POST"])
def analyze_single():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json
    model = data.get("model", "gpt-4o")
    lang = data.get("lang", "cs")

    email_text = (
        f"Od: {data.get('from', '?')}\n"
        f"Předmět: {data.get('subject', '?')}\n"
        f"Datum: {data.get('date', '?')}\n"
        f"Tělo:\n{data.get('body', '')}"
    )

    lang_instruction = ""
    if lang == "en":
        lang_instruction = "\nPiš odpovědi v angličtině."
    elif lang == "auto":
        lang_instruction = "\nPiš odpovědi ve stejném jazyce jako je e-mail."

    try:
        response = client.chat.completions.create(
            model=_resolve_model(model),
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT + lang_instruction},
                {"role": "user", "content": email_text}
            ],
            temperature=0.3,
            max_tokens=1000
        )
        result = json.loads(response.choices[0].message.content)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/analyze-inbox", methods=["POST"])
def analyze_inbox():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json
    token = data.get("token")
    rest_url = data.get("restUrl", "https://graph.microsoft.com/v1.0")
    top = data.get("top", 20)
    model = data.get("model", "gpt-4o")
    lang = data.get("lang", "cs")

    # Fetch emails via Graph API using the callback token from Office.js
    # restUrl z Office.js může být outlook.office.com REST URL — převedeme na Graph
    graph_url = "https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages"
    params = {
        "$select": "id,subject,from,receivedDateTime,bodyPreview,isRead,importance",
        "$top": top,
        "$orderby": "receivedDateTime DESC"
    }

    try:
        resp = http_requests.get(
            graph_url,
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=15
        )
        resp.raise_for_status()
        messages = resp.json().get("value", [])
    except Exception as e:
        return jsonify({"error": f"Graph API chyba: {str(e)}"}), 502

    # Build text for AI
    email_texts = []
    for msg in messages:
        sender = "?"
        if msg.get("from", {}).get("emailAddress"):
            ea = msg["from"]["emailAddress"]
            sender = f"{ea.get('name', '')} <{ea.get('address', '')}>"
        email_texts.append(
            f"ID: {msg['id'][:20]}\n"
            f"Od: {sender}\n"
            f"Předmět: {msg.get('subject', '(bez předmětu)')}\n"
            f"Datum: {msg.get('receivedDateTime', '')}\n"
            f"Důležitost: {msg.get('importance', 'normal')}\n"
            f"Přečteno: {msg.get('isRead', False)}\n"
            f"Náhled: {msg.get('bodyPreview', '')[:300]}"
        )

    combined = "\n---\n".join(email_texts)

    try:
        response = client.chat.completions.create(
            model=_resolve_model(model),
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": INBOX_SYSTEM_PROMPT},
                {"role": "user", "content": combined}
            ],
            temperature=0.3,
            max_tokens=4000
        )
        result = json.loads(response.choices[0].message.content)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def _resolve_model(model_key):
    """Map frontend model names to API model IDs."""
    mapping = {
        "gpt-4o": "gpt-4o",
        "gpt-4o-mini": "gpt-4o-mini",
        "claude-sonnet": "gpt-4o",  # fallback — pro Claude použijte Anthropic SDK
        "local": "gpt-4o",          # fallback — pro Ollama změňte base_url
    }
    return mapping.get(model_key, "gpt-4o")


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "version": "1.0.0"})


if __name__ == "__main__":
    app.run(debug=True, port=5000)

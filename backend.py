import os
import json
import base64
from flask import Flask, request, jsonify, send_from_directory, Response
from flask_cors import CORS
from openai import OpenAI
import requests

app = Flask(__name__)
CORS(app)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TRANSPARENT_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9l9oQAAAAASUVORK5CYII="
)

_client = None


def get_client():
    global _client
    if _client is not None:
        return _client

    api_key = os.environ.get("EINFRA_API_KEY") or os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "Missing API key: set EINFRA_API_KEY (preferred) or OPENAI_API_KEY."
        )

    _client = OpenAI(
        api_key=api_key,
        base_url=os.environ.get("EINFRA_BASE_URL", "https://llm.ai.e-infra.cz/v1/"),
    )
    return _client

API_SECRET = os.environ.get("API_SECRET", "")
DEFAULT_MODEL = os.environ.get("EINFRA_MODEL", "gpt-4o-mini")

DEFAULT_PROMPT = """Jsi e-mailový asistent. Zpracuj e-mail a vrať JSON s poli:
priority: P0_urgent|P1_action_needed|P2_informational|P3_ignorable
category: administrativa|výuka|IT|výzkum|osobní|spam|newsletter
summary: krátký souhrn česky
needs_reply: true/false
reply_deadline: dnes|tento_tyden|zadny
suggested_reply: návrh odpovědi nebo null
Zohledni custom pravidla uživatele."""

INBOX_PROMPT = """Jsi e-mailový asistent. Zpracuj seznam e-mailů a vrať JSON {"emails": [...]} se stejnými poli jako u jednoho e-mailu. Řaď od nejvyšší priority."""


def check_auth():
    if not API_SECRET:
        return True
    return request.headers.get("Authorization", "") == f"Bearer {API_SECRET}"


def get_prompt(data):
    return data.get("customPrompt") or DEFAULT_PROMPT


def get_priority_senders(data):
    return data.get("prioritySenders", [])


def merge_prompt(custom_prompt, senders):
    extra = []
    if senders:
        extra.append("Preferovaní odesílatelé: " + ", ".join(senders))
    if extra:
        return custom_prompt + "\n\n" + "\n".join(extra)
    return custom_prompt


@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "taskpane.html")


@app.route("/taskpane.html")
def taskpane():
    return send_from_directory(BASE_DIR, "taskpane.html")


@app.route("/assets/<path:filename>")
def assets(filename):
    assets_dir = os.path.join(BASE_DIR, "assets")
    asset_path = os.path.join(assets_dir, filename)

    if os.path.isfile(asset_path):
        return send_from_directory(assets_dir, filename)

    if filename in {"icon-64.png", "icon-128.png"}:
        return Response(TRANSPARENT_PNG, mimetype="image/png")

    return jsonify({"error": "Asset not found"}), 404


@app.route("/analyze", methods=["POST"])
def analyze():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401
    data = request.json or {}
    senders = get_priority_senders(data)
    prompt = merge_prompt(get_prompt(data), senders)
    email = (
        f"Od: {data.get('from', '?')}\n"
        f"Předmět: {data.get('subject', '?')}\n"
        f"Datum: {data.get('date', '?')}\n"
        "Tělo:\n"
        f"{data.get('body', '')}\n\n"
        f"Preferovaní odesílatelé: {', '.join(senders) if senders else 'žádní'}"
    )
    try:
        client = get_client()
        resp = client.chat.completions.create(
            model=data.get("model", DEFAULT_MODEL),
            response_format={"type": "json_object"},
            messages=[{"role":"system","content": prompt}, {"role":"user","content": email}],
            temperature=0.3,
            max_tokens=1000,
        )
        return jsonify(json.loads(resp.choices[0].message.content))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/analyze-inbox", methods=["POST"])
def analyze_inbox():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401
    data = request.json or {}
    token = data.get("token")
    top = int(data.get("top", 20))
    prompt = merge_prompt(data.get("customPrompt") or INBOX_PROMPT, data.get("prioritySenders") or [])
    priority_senders = [s.lower() for s in (data.get("prioritySenders") or [])]

    r = requests.get(
        "https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages",
        headers={"Authorization": f"Bearer {token}"},
        params={"$select":"id,subject,from,receivedDateTime,bodyPreview,isRead,importance", "$top": top, "$orderby":"receivedDateTime DESC"},
        timeout=20,
    )
    r.raise_for_status()
    msgs = r.json().get("value", [])

    items = []
    for m in msgs:
        ea = m.get("from", {}).get("emailAddress", {})
        sender = f"{ea.get('name','')} <{ea.get('address','')}>".strip()
        items.append({
            "id": m.get("id"),
            "subject": m.get("subject", "(bez předmětu)"),
            "from": sender,
            "receivedDateTime": m.get("receivedDateTime", ""),
            "bodyPreview": m.get("bodyPreview", ""),
            "isRead": m.get("isRead", False),
            "importance": m.get("importance", "normal"),
            "priorityBoost": any(s in sender.lower() for s in priority_senders),
        })

    user_text = json.dumps(items, ensure_ascii=False)
    try:
        client = get_client()
        resp = client.chat.completions.create(
            model=data.get("model", DEFAULT_MODEL),
            response_format={"type": "json_object"},
            messages=[{"role":"system","content": prompt}, {"role":"user","content": user_text}],
            temperature=0.3,
            max_tokens=3000,
        )
        result = json.loads(resp.choices[0].message.content)
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True, port=5000)

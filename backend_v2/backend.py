import base64
import json
import os
import re
from datetime import datetime, timedelta, timezone

import requests
from flask import Flask, Response, jsonify, request
from flask_cors import CORS
from openai import OpenAI


app = Flask(__name__)
CORS(app)

API_SECRET = os.environ.get("API_SECRET", "")
DEFAULT_MODEL = os.environ.get("EINFRA_MODEL", "gpt-4o-mini")
DEFAULT_BASE_URL = os.environ.get("EINFRA_BASE_URL", "https://llm.ai.e-infra.cz/v1/")
MAX_EMAILS_FOR_LLM = 120

TRANSPARENT_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9l9oQAAAAASUVORK5CYII="
)

MAILAI_CATEGORY_PREFIX = "MailAI/"
MAILAI_CATEGORY_MAP = {
    "urgentni": ("MailAI/Urgentni", "preset0"),
    "stredne_dulezite": ("MailAI/Stredne dulezite", "preset1"),
    "pocka": ("MailAI/Pocka", "preset3"),
    "k_preposlani": ("MailAI/K preposlani", "preset7"),
    "ignorovat": ("MailAI/Ignorovat", "preset12"),
}
MAILAI_DEADLINE_CATEGORY = ("MailAI/S terminem", "preset6")
BUCKET_ORDER = tuple(MAILAI_CATEGORY_MAP.keys())

DEFAULT_PROMPT = """Jsi e-mailovy asistent. Zpracuj e-mail a vrat JSON s poli:
priority: P0_urgent|P1_action_needed|P2_informational|P3_ignorable
category: administrativa|vyuka|IT|vyzkum|osobni|spam|newsletter
summary: kratky souhrn cesky
needs_reply: true/false
reply_deadline: dnes|tento_tyden|zadny
suggested_reply: navrh odpovedi nebo null
Zohledni custom pravidla uzivatele."""

INBOX_PROMPT = """Jsi e-mailovy asistent. Zpracuj seznam neprectenych e-mailu a vrat pouze validni JSON s touto strukturou:
{
  "overview": "strucny souhrn",
  "counts": {
    "urgentni": number,
    "stredne_dulezite": number,
    "pocka": number,
    "k_preposlani": number,
    "ignorovat": number
  },
  "buckets": {
    "urgentni": [{"id":"...","subject":"...","from":"...","reason":"...","action":"...","has_deadline":true|false,"deadline_hint":"..."}],
    "stredne_dulezite": [{"id":"...","subject":"...","from":"...","reason":"...","action":"...","has_deadline":true|false,"deadline_hint":"..."}],
    "pocka": [{"id":"...","subject":"...","from":"...","reason":"...","action":"..."}],
    "k_preposlani": [{"id":"...","subject":"...","from":"...","reason":"...","forward_to":"...","action":"..."}],
    "ignorovat": [{"id":"...","subject":"...","from":"...","reason":"...","action":"oznacit_jako_prectene|ignorovat"}]
  },
  "recommended_bulk_actions": {
    "mark_read_ids": ["..."]
  }
}
Pravidla:
- Pouzij jen kategorie: urgentni, stredne_dulezite, pocka, k_preposlani, ignorovat.
- Kazdy e-mail zarad prave do jedne kategorie.
- Nikdy nenavrhuj mazani e-mailu ani akci smazat.
- U urgentni a stredne_dulezite: has_deadline=true pokud e-mail obsahuje konkretni termin/datum. deadline_hint dej strucne cesky nebo null.
- Odpovidej cesky.
"""


def check_auth() -> bool:
    if not API_SECRET:
        return True
    return request.headers.get("Authorization", "") == f"Bearer {API_SECRET}"


def get_client(data: dict) -> OpenAI:
    api_key = (data or {}).get("llmApiKey")
    if not api_key:
        raise RuntimeError("Missing LLM API key in request (llmApiKey).")

    return OpenAI(
        api_key=api_key,
        base_url=(data or {}).get("llmBaseUrl") or DEFAULT_BASE_URL,
        timeout=int((data or {}).get("llmTimeout", 60)),
        max_retries=1,
    )


def parse_json_content(content: str) -> dict:
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        cleaned = (content or "").replace("```json", "").replace("```", "").strip()
        return json.loads(cleaned)


def merge_prompt(base_prompt: str, custom_prompt: str, senders: list[str]) -> str:
    parts = [(base_prompt or "").strip()]
    cp = (custom_prompt or "").strip()
    if cp:
        parts.append("Doplnujici instrukce uzivatele:\n" + cp)
    if senders:
        parts.append("Preferovani odesilatele: " + ", ".join(senders))
    return "\n\n".join([p for p in parts if p])


def enforce_no_delete_policy(result: dict) -> dict:
    actions = result.get("recommended_bulk_actions") or {}
    actions["delete_ids"] = []
    result["recommended_bulk_actions"] = actions

    buckets = result.get("buckets") or {}
    ignore_bucket = buckets.get("ignorovat") or []
    for item in ignore_bucket:
        if item.get("action") == "smazat":
            item["action"] = "ignorovat"
    buckets["ignorovat"] = ignore_bucket
    result["buckets"] = buckets
    return result


def _fetch_graph_messages(token: str, endpoint: str, query: dict, max_items: int) -> list[dict]:
    items = []
    url = endpoint
    params = query.copy()

    while url and len(items) < max_items:
        r = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=30,
        )
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None

    return items[:max_items]


def _normalize_inbox_items(messages: list[dict]) -> list[dict]:
    items = []
    for m in messages:
        ea = m.get("from", {}).get("emailAddress", {})
        sender = f"{ea.get('name', '')} <{ea.get('address', '')}>".strip()
        items.append(
            {
                "id": m.get("id"),
                "conversationId": m.get("conversationId"),
                "subject": m.get("subject", "(bez predmetu)"),
                "from": sender,
                "receivedDateTime": m.get("receivedDateTime", ""),
                "bodyPreview": m.get("bodyPreview", ""),
                "importance": m.get("importance", "normal"),
                "categories": m.get("categories", []),
                "webLink": m.get("webLink", ""),
                "isRead": m.get("isRead", False),
            }
        )
    return items


def _parse_graph_datetime(value: str) -> datetime:
    try:
        return datetime.fromisoformat((value or "").replace("Z", "+00:00"))
    except Exception:
        return datetime.now(timezone.utc)


def fetch_unread_messages(token: str, days: int, top: int) -> list[dict]:
    since = (
        (datetime.now(timezone.utc) - timedelta(days=days))
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )
    messages = _fetch_graph_messages(
        token,
        "https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages",
        {
            "$select": "id,conversationId,subject,from,receivedDateTime,bodyPreview,isRead,importance,categories,webLink",
            "$top": top,
            "$orderby": "receivedDateTime DESC",
            "$filter": f"isRead eq false and receivedDateTime ge {since}",
        },
        top,
    )
    return _normalize_inbox_items(messages)


def fetch_not_replied_messages(token: str, days: int, top: int) -> list[dict]:
    since = (
        (datetime.now(timezone.utc) - timedelta(days=days))
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )

    inbox_messages = _fetch_graph_messages(
        token,
        "https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages",
        {
            "$select": "id,conversationId,subject,from,receivedDateTime,bodyPreview,isRead,importance,categories,webLink",
            "$top": min(top, 200),
            "$orderby": "receivedDateTime DESC",
            "$filter": f"receivedDateTime ge {since}",
        },
        top,
    )

    sent_messages = _fetch_graph_messages(
        token,
        "https://graph.microsoft.com/v1.0/me/mailfolders/sentitems/messages",
        {
            "$select": "conversationId,createdDateTime",
            "$top": min(max(top * 2, 200), 500),
            "$orderby": "createdDateTime DESC",
            "$filter": f"createdDateTime ge {since}",
        },
        max(top * 2, 200),
    )

    replied_conversations = {m.get("conversationId") for m in sent_messages if m.get("conversationId")}
    not_replied = [m for m in inbox_messages if m.get("conversationId") not in replied_conversations]
    return _normalize_inbox_items(not_replied[:top])


def filter_items_for_analysis(items: list[dict], mode: str, urgent_reminder_hours: int) -> tuple[list[dict], dict]:
    if mode == "Vsechny (vcetne jiz oznacenych)":
        return items, {"skipped_labeled": 0, "urgent_reincluded": 0}

    unlabeled = []
    labeled = []
    for item in items:
        categories = item.get("categories") or []
        if any(str(cat).startswith(MAILAI_CATEGORY_PREFIX) for cat in categories):
            labeled.append(item)
        else:
            unlabeled.append(item)

    if mode == "Jen bez MailAI stitku":
        return unlabeled, {"skipped_labeled": len(labeled), "urgent_reincluded": 0}

    cutoff = datetime.now(timezone.utc) - timedelta(hours=max(1, urgent_reminder_hours))
    urgent_reminders = []
    for item in labeled:
        categories = set(item.get("categories") or [])
        if "MailAI/Urgentni" not in categories:
            continue
        received = _parse_graph_datetime(item.get("receivedDateTime", ""))
        if received <= cutoff:
            urgent_reminders.append(item)

    seen_ids = {item.get("id") for item in unlabeled}
    for item in urgent_reminders:
        msg_id = item.get("id")
        if msg_id and msg_id not in seen_ids:
            unlabeled.append(item)
            seen_ids.add(msg_id)

    return unlabeled, {"skipped_labeled": len(labeled), "urgent_reincluded": len(urgent_reminders)}


def summarize_unread(client: OpenAI, model: str, prompt: str, items: list[dict], days: int) -> dict:
    prepared_items = []
    for itm in items:
        prepared_items.append(
            {
                "id": str(itm.get("id") or ""),
                "conversationId": str(itm.get("conversationId") or ""),
                "subject": str(itm.get("subject") or "(bez predmetu)"),
                "from": str(itm.get("from") or "(neznamy odesilatel)"),
                "receivedDateTime": str(itm.get("receivedDateTime") or ""),
                "bodyPreview": str(itm.get("bodyPreview") or ""),
                "importance": str(itm.get("importance") or "normal"),
                "categories": [str(c) for c in (itm.get("categories") or [])],
            }
        )

    user_text = json.dumps(
        {
            "window_days": days,
            "total_unread_fetched": len(prepared_items),
            "emails": prepared_items,
        },
        ensure_ascii=False,
    )

    try:
        resp = client.chat.completions.create(
            model=model,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": prompt},
                {"role": "user", "content": user_text},
            ],
            temperature=0.2,
            max_tokens=3500,
        )
        return parse_json_content(resp.choices[0].message.content)
    except Exception:
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": prompt + "\n\nVratis pouze validni JSON bez markdownu."},
                {"role": "user", "content": user_text},
            ],
            temperature=0.2,
            max_tokens=3500,
        )
        return parse_json_content(resp.choices[0].message.content)


def enrich_result_with_source_metadata(result: dict, source_items: list[dict]) -> dict:
    source_by_id = {str(item.get("id") or ""): item for item in source_items if item.get("id")}
    buckets = result.get("buckets") or {}

    for bucket_key in BUCKET_ORDER:
        enriched_items = []
        for itm in buckets.get(bucket_key, []):
            msg_id = str(itm.get("id") or "")
            source = source_by_id.get(msg_id, {})
            merged = dict(itm)
            if source.get("receivedDateTime"):
                merged["receivedDateTime"] = source.get("receivedDateTime")
            if source.get("webLink"):
                merged["webLink"] = source.get("webLink")
            if source.get("categories"):
                merged["categories"] = source.get("categories")
            enriched_items.append(merged)
        buckets[bucket_key] = enriched_items

    result["buckets"] = buckets
    return result


def graph_get_master_categories(token: str) -> set[str]:
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
        headers={"Authorization": f"Bearer {token}"},
        timeout=20,
    )
    r.raise_for_status()
    return {c.get("displayName", "") for c in r.json().get("value", []) if c.get("displayName")}


def graph_create_master_category(token: str, name: str, color: str) -> None:
    r = requests.post(
        "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"displayName": name, "color": color},
        timeout=20,
    )
    if r.status_code not in (200, 201, 409):
        r.raise_for_status()


def ensure_selected_master_categories(
    token: str,
    bucket_label_map: dict[str, str],
    add_deadline_label: bool,
    deadline_label_name: str,
) -> None:
    existing = graph_get_master_categories(token)
    for bucket_key in BUCKET_ORDER:
        label_name = str(bucket_label_map.get(bucket_key, "") or "").strip()
        if not label_name:
            continue
        color = MAILAI_CATEGORY_MAP[bucket_key][1]
        if label_name not in existing:
            graph_create_master_category(token, label_name, color)
            existing.add(label_name)

    dl_name = (deadline_label_name or "").strip()
    if add_deadline_label and dl_name and dl_name not in existing:
        graph_create_master_category(token, dl_name, MAILAI_DEADLINE_CATEGORY[1])


def graph_assign_category(token: str, msg_id: str, category_name: str) -> None:
    get_r = requests.get(
        f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}",
        headers={"Authorization": f"Bearer {token}"},
        params={"$select": "categories"},
        timeout=20,
    )
    get_r.raise_for_status()
    current_categories = get_r.json().get("categories", [])
    if category_name in current_categories:
        return

    patch_r = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"categories": current_categories + [category_name]},
        timeout=20,
    )
    patch_r.raise_for_status()


def graph_patch_read(token: str, msg_id: str) -> None:
    r = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"isRead": True},
        timeout=20,
    )
    r.raise_for_status()


def graph_create_calendar_event(
    token: str,
    subject: str,
    start_dt: datetime,
    end_dt: datetime,
    timezone_name: str,
    body_text: str,
    categories: list[str] | None = None,
) -> dict:
    payload = {
        "subject": subject,
        "body": {"contentType": "Text", "content": body_text},
        "start": {"dateTime": start_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": timezone_name},
        "end": {"dateTime": end_dt.strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": timezone_name},
        "isReminderOn": True,
        "categories": categories or [MAILAI_DEADLINE_CATEGORY[0]],
    }
    r = requests.post(
        "https://graph.microsoft.com/v1.0/me/events",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
        timeout=20,
    )
    r.raise_for_status()
    return r.json()


def decode_jwt_claims_unverified(token: str) -> dict:
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return {}
        payload = parts[1]
        padding = "=" * (-len(payload) % 4)
        decoded = base64.urlsafe_b64decode(payload + padding)
        return json.loads(decoded.decode("utf-8"))
    except Exception:
        return {}


def graph_endpoint_status(token: str, url: str, params: dict | None = None) -> tuple[int, str]:
    try:
        r = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params=params,
            timeout=20,
        )
        if r.status_code < 400:
            return r.status_code, "OK"
        try:
            detail = r.json().get("error", {}).get("message") or r.text[:200]
        except Exception:
            detail = r.text[:200]
        return r.status_code, detail
    except Exception as e:
        return 0, str(e)


def parse_deadline_date_hint(hint: str | None) -> str | None:
    text = (hint or "").strip().lower()
    if not text:
        return None

    now_local = datetime.now().astimezone()

    if "dnes" in text:
        return now_local.date().isoformat()
    if "zitra" in text or "zitra" in text:
        return (now_local + timedelta(days=1)).date().isoformat()

    m = re.search(r"\b(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{2,4})\b", text)
    if m:
        day = int(m.group(1))
        month = int(m.group(2))
        year = int(m.group(3))
        if year < 100:
            year += 2000
        try:
            return now_local.replace(year=year, month=month, day=day).date().isoformat()
        except ValueError:
            return None

    m = re.search(r"\b(\d{1,2})\s*[./-]\s*(\d{1,2})(?:\s|$)", text)
    if m:
        day = int(m.group(1))
        month = int(m.group(2))
        year = now_local.year
        try:
            candidate = now_local.replace(year=year, month=month, day=day)
            if candidate.date() < now_local.date():
                candidate = candidate.replace(year=year + 1)
            return candidate.date().isoformat()
        except ValueError:
            return None

    return None


def get_priority_senders(data: dict) -> list[str]:
    senders = data.get("prioritySenders", [])
    if isinstance(senders, str):
        senders = [s.strip() for s in re.split(r"\r?\n|,", senders) if s.strip()]
    if not isinstance(senders, list):
        return []
    return [str(s).strip() for s in senders if str(s).strip()]


@app.route("/", methods=["GET"])
def index():
    return jsonify(
        {
            "service": "MailAI backend_v2",
            "status": "ok",
            "endpoints": [
                "/health",
                "/models",
                "/analyze",
                "/analyze-inbox",
                "/graph/diagnostics",
                "/graph/categories",
                "/apply-classification",
                "/calendar/create-event",
            ],
        }
    )


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"}), 200


@app.route("/assets/<path:filename>", methods=["GET"])
def assets(filename: str):
    if filename in {"icon-64.png", "icon-128.png"}:
        return Response(TRANSPARENT_PNG, mimetype="image/png")
    return jsonify({"error": "Asset not found"}), 404


@app.route("/models", methods=["POST"])
def models():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json or {}
    try:
        client = get_client(data)
        resp = client.models.list()
        models_list = sorted([m.id for m in getattr(resp, "data", []) if getattr(m, "id", None)])
        return jsonify({"models": models_list, "defaultModel": DEFAULT_MODEL})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/analyze", methods=["POST"])
def analyze_single_email():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json or {}
    senders = get_priority_senders(data)
    prompt = merge_prompt(data.get("customPrompt") or DEFAULT_PROMPT, "", senders)
    email = (
        f"Od: {data.get('from', '?')}\n"
        f"Predmet: {data.get('subject', '?')}\n"
        f"Datum: {data.get('date', '?')}\n"
        "Telo:\n"
        f"{data.get('body', '')}\n\n"
        f"Preferovani odesilatele: {', '.join(senders) if senders else 'zadni'}"
    )

    try:
        client = get_client(data)
        resp = client.chat.completions.create(
            model=data.get("model") or DEFAULT_MODEL,
            response_format={"type": "json_object"},
            messages=[{"role": "system", "content": prompt}, {"role": "user", "content": email}],
            temperature=0.3,
            max_tokens=1200,
        )
        return jsonify(parse_json_content(resp.choices[0].message.content))
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/analyze-inbox", methods=["POST"])
def analyze_inbox():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json or {}
    token = data.get("token")
    if not token:
        return jsonify({"error": "Missing Graph token in request (token)."}), 400

    top = int(data.get("top", 200))
    days = int(data.get("days", 10))
    analysis_mode = data.get("analysisMode", "Neprectene")
    label_handling_mode = data.get("labelHandlingMode", "Jen bez MailAI stitku + urgentni pripomenuti")
    urgent_reminder_hours = int(data.get("urgentReminderHours", 24))
    llm_limit = int(data.get("maxEmailsForLlm", MAX_EMAILS_FOR_LLM))
    llm_limit = max(10, min(llm_limit, 500))

    senders = get_priority_senders(data)
    prompt = merge_prompt(INBOX_PROMPT, data.get("customPrompt", ""), senders)

    try:
        if analysis_mode == "Bez odpovedi (Inbox vs Sent)":
            items = fetch_not_replied_messages(token, days, top)
        else:
            items = fetch_unread_messages(token, days, top)

        fetched_count = len(items)
        filtered_items, filter_stats = filter_items_for_analysis(items, label_handling_mode, urgent_reminder_hours)

        if not filtered_items:
            empty = {
                "overview": "Po filtraci nejsou zadne e-maily k analyze.",
                "counts": {k: 0 for k in BUCKET_ORDER},
                "buckets": {k: [] for k in BUCKET_ORDER},
                "recommended_bulk_actions": {"mark_read_ids": []},
            }
            return jsonify(
                {
                    "result": empty,
                    "meta": {
                        "fetched_count": fetched_count,
                        "analyzed_count": 0,
                        "skipped_labeled": filter_stats["skipped_labeled"],
                        "urgent_reincluded": filter_stats["urgent_reincluded"],
                        "analysis_mode": analysis_mode,
                        "label_handling_mode": label_handling_mode,
                    },
                }
            )

        llm_items = filtered_items[:llm_limit]
        truncated = len(filtered_items) > len(llm_items)

        client = get_client(data)
        result = summarize_unread(
            client,
            data.get("model") or DEFAULT_MODEL,
            prompt,
            llm_items,
            days,
        )
        result = enforce_no_delete_policy(result)
        result = enrich_result_with_source_metadata(result, filtered_items)

        for bucket_key in BUCKET_ORDER:
            bucket_items = (result.get("buckets") or {}).get(bucket_key, [])
            (result.get("counts") or {})[bucket_key] = len(bucket_items)
            for itm in bucket_items:
                if itm.get("has_deadline") and not itm.get("deadline_suggested_date"):
                    itm["deadline_suggested_date"] = parse_deadline_date_hint(itm.get("deadline_hint"))

        return jsonify(
            {
                "result": result,
                "meta": {
                    "fetched_count": fetched_count,
                    "analyzed_count": len(llm_items),
                    "skipped_labeled": filter_stats["skipped_labeled"],
                    "urgent_reincluded": filter_stats["urgent_reincluded"],
                    "analysis_mode": analysis_mode,
                    "label_handling_mode": label_handling_mode,
                    "truncated_for_llm": truncated,
                    "llm_limit": llm_limit,
                },
            }
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/graph/diagnostics", methods=["POST"])
def graph_diagnostics():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json or {}
    token = data.get("token")
    if not token:
        return jsonify({"error": "Missing Graph token in request (token)."}), 400

    claims = decode_jwt_claims_unverified(token)
    scopes = str(claims.get("scp", ""))
    scope_set = set(scopes.split()) if scopes else set()

    me_status, me_msg = graph_endpoint_status(token, "https://graph.microsoft.com/v1.0/me", {"$select": "id,userPrincipalName"})
    msg_status, msg_msg = graph_endpoint_status(token, "https://graph.microsoft.com/v1.0/me/messages", {"$top": 1, "$select": "id"})
    cat_status, cat_msg = graph_endpoint_status(token, "https://graph.microsoft.com/v1.0/me/outlook/masterCategories")
    evt_status, evt_msg = graph_endpoint_status(token, "https://graph.microsoft.com/v1.0/me/events", {"$top": 1, "$select": "id"})

    return jsonify(
        {
            "claims": {
                "aud": claims.get("aud"),
                "scp": claims.get("scp"),
                "roles": claims.get("roles", []),
            },
            "endpoints": {
                "/me": {"status": me_status, "detail": me_msg},
                "/me/messages": {"status": msg_status, "detail": msg_msg},
                "/me/outlook/masterCategories": {"status": cat_status, "detail": cat_msg},
                "/me/events": {"status": evt_status, "detail": evt_msg},
            },
            "scope_checks": {
                "Mail.Read": "Mail.Read" in scope_set,
                "Mail.ReadWrite": "Mail.ReadWrite" in scope_set,
                "MailboxSettings.ReadWrite": "MailboxSettings.ReadWrite" in scope_set,
                "Calendars.ReadWrite": "Calendars.ReadWrite" in scope_set,
            },
        }
    )


@app.route("/graph/categories", methods=["POST"])
def graph_categories():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json or {}
    token = data.get("token")
    if not token:
        return jsonify({"error": "Missing Graph token in request (token)."}), 400

    try:
        categories = sorted(graph_get_master_categories(token))
        return jsonify({"categories": categories, "count": len(categories)})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/apply-classification", methods=["POST"])
def apply_classification():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json or {}
    token = data.get("token")
    result = data.get("result") or {}
    if not token:
        return jsonify({"error": "Missing Graph token in request (token)."}), 400

    use_custom_label_mapping = bool(data.get("useCustomLabelMapping", False))
    add_deadline_label = bool(data.get("addDeadlineLabel", True))
    deadline_label_name = str(data.get("deadlineLabelName") or MAILAI_DEADLINE_CATEGORY[0]).strip()
    create_categories = bool(data.get("createMissingCategories", True))

    if use_custom_label_mapping:
        raw_map = data.get("bucketLabelMap") or {}
        bucket_label_map = {}
        for bucket_key in BUCKET_ORDER:
            value = str(raw_map.get(bucket_key, "") or "").strip()
            bucket_label_map[bucket_key] = value or MAILAI_CATEGORY_MAP[bucket_key][0]
    else:
        bucket_label_map = {bucket_key: MAILAI_CATEGORY_MAP[bucket_key][0] for bucket_key in BUCKET_ORDER}

    try:
        if create_categories:
            ensure_selected_master_categories(token, bucket_label_map, add_deadline_label, deadline_label_name)

        ok = 0
        fail = 0
        deadline_ok = 0
        deadline_fail = 0

        buckets = result.get("buckets") or {}
        for bucket_key in BUCKET_ORDER:
            category_name = bucket_label_map.get(bucket_key, "").strip()
            if not category_name:
                continue
            for itm in buckets.get(bucket_key, []):
                msg_id = str(itm.get("id") or "").strip()
                if not msg_id:
                    continue
                try:
                    graph_assign_category(token, msg_id, category_name)
                    ok += 1
                except Exception:
                    fail += 1

                if add_deadline_label and deadline_label_name and bucket_key in ("urgentni", "stredne_dulezite") and itm.get("has_deadline"):
                    try:
                        graph_assign_category(token, msg_id, deadline_label_name)
                        deadline_ok += 1
                    except Exception:
                        deadline_fail += 1

        mark_ids = data.get("markReadIds")
        if not isinstance(mark_ids, list):
            mark_ids = ((result.get("recommended_bulk_actions") or {}).get("mark_read_ids") or [])

        mark_ok = 0
        mark_fail = 0
        for msg_id in mark_ids:
            raw_id = str(msg_id or "").strip()
            if not raw_id:
                continue
            try:
                graph_patch_read(token, raw_id)
                mark_ok += 1
            except Exception:
                mark_fail += 1

        return jsonify(
            {
                "assigned_ok": ok,
                "assigned_fail": fail,
                "deadline_assigned_ok": deadline_ok,
                "deadline_assigned_fail": deadline_fail,
                "mark_read_ok": mark_ok,
                "mark_read_fail": mark_fail,
            }
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/calendar/create-event", methods=["POST"])
def calendar_create_event():
    if not check_auth():
        return jsonify({"error": "Unauthorized"}), 401

    data = request.json or {}
    token = data.get("token")
    if not token:
        return jsonify({"error": "Missing Graph token in request (token)."}), 400

    subject = str(data.get("subject") or "MailAI termin")
    start_iso = data.get("startDateTime")
    end_iso = data.get("endDateTime")
    timezone_name = str(data.get("timeZone") or "Europe/Prague")
    body_text = str(data.get("bodyText") or "Vytvoreno pres MailAI backend_v2")
    categories = data.get("categories") if isinstance(data.get("categories"), list) else None

    if not start_iso or not end_iso:
        return jsonify({"error": "startDateTime and endDateTime are required."}), 400

    try:
        start_dt = datetime.fromisoformat(str(start_iso).replace("Z", "+00:00"))
        end_dt = datetime.fromisoformat(str(end_iso).replace("Z", "+00:00"))
        created = graph_create_calendar_event(token, subject, start_dt, end_dt, timezone_name, body_text, categories)
        return jsonify({"event": created})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5001)

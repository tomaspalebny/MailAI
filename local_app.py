import json
import base64
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
import streamlit as st
from openai import OpenAI
from openai import APITimeoutError


INBOX_PROMPT = """Jsi e-mailový asistent. Zpracuj seznam nepřečtených e-mailů a vrať pouze validní JSON s touto strukturou:
{
  "overview": "stručný souhrn",
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
- Použij jen kategorie: urgentni, stredne_dulezite, pocka, k_preposlani, ignorovat.
- Každý e-mail zařaď právě do jedné kategorie.
- Nikdy nenavrhuj mazání e-mailů ani akci smazat.
- U urgentni a stredne_dulezite: has_deadline=true pokud e-mail obsahuje konkrétní termín/datum (deadline, uzávěrka, schůzka, do kdy). deadline_hint = stručný popis termínu česky, nebo null.
- Odpovídej česky.
"""

MAX_EMAILS_FOR_LLM = 120
SETTINGS_FILE = Path(".mailai_local_settings.json")
MAILAI_CATEGORY_PREFIX = "MailAI/"
# preset0=červená, preset1=oranžová, preset3=žlutá, preset7=modrá, preset12=šedá, preset6=fialová
MAILAI_CATEGORY_MAP = {
    "urgentni":          ("MailAI/Urgentni",          "preset0"),
    "stredne_dulezite": ("MailAI/Stredne dulezite",  "preset1"),
    "pocka":            ("MailAI/Pocka",              "preset3"),
    "k_preposlani":     ("MailAI/K preposlani",       "preset7"),
    "ignorovat":        ("MailAI/Ignorovat",          "preset12"),
}
MAILAI_DEADLINE_CATEGORY = ("MailAI/S terminem", "preset6")  # fialová

# Barvy pro Streamlit UI  (hex, odpovídají Outlook presetům výše)
BUCKET_UI = {
    "urgentni":         {"label": "Urgentní",          "color": "#e74c3c", "emoji": "🔴"},
    "stredne_dulezite": {"label": "Středně důležité",  "color": "#e67e22", "emoji": "🟠"},
    "pocka":            {"label": "Počká",             "color": "#d4ac0d", "emoji": "🟡"},
    "k_preposlani":     {"label": "K přeposlání",      "color": "#2980b9", "emoji": "🔵"},
    "ignorovat":        {"label": "Ignorovat",         "color": "#95a5a6", "emoji": "⚫"},
}
BUCKET_ORDER = tuple(BUCKET_UI.keys())


def load_local_settings() -> dict:
    if not SETTINGS_FILE.exists():
        return {}
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_local_settings(settings: dict) -> None:
    SETTINGS_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def build_settings_payload() -> dict:
    return {
        "llm_api_key": st.session_state.get("llm_api_key", ""),
        "llm_base_url": st.session_state.get("llm_base_url", "https://llm.ai.e-infra.cz/v1/"),
        "llm_timeout": int(st.session_state.get("llm_timeout", 60)),
        "analysis_mode": st.session_state.get("analysis_mode", "Nepřečtené"),
        "label_handling_mode": st.session_state.get(
            "label_handling_mode", "Jen bez MailAI štítku + urgentní připomenutí"
        ),
        "urgent_reminder_hours": int(st.session_state.get("urgent_reminder_hours", 24)),
        "model": st.session_state.get("model", ""),
        "graph_token": st.session_state.get("graph_token_input", ""),
        "days": int(st.session_state.get("days", 10)),
        "top": int(st.session_state.get("top", 200)),
        "custom_prompt": st.session_state.get("custom_prompt", ""),
        "priority_senders_raw": st.session_state.get("priority_senders_raw", ""),
        "calendar_timezone": st.session_state.get("calendar_timezone", "Europe/Prague"),
        "auto_save_settings": bool(st.session_state.get("auto_save_settings", True)),
    }


def initialize_state_from_settings() -> None:
    if st.session_state.get("settings_initialized"):
        return
    saved = load_local_settings()
    st.session_state["llm_api_key"] = saved.get("llm_api_key", "")
    st.session_state["llm_base_url"] = saved.get("llm_base_url", "https://llm.ai.e-infra.cz/v1/")
    st.session_state["llm_timeout"] = int(saved.get("llm_timeout", 60))
    st.session_state["analysis_mode"] = saved.get("analysis_mode", "Nepřečtené")
    st.session_state["label_handling_mode"] = saved.get(
        "label_handling_mode", "Jen bez MailAI štítku + urgentní připomenutí"
    )
    st.session_state["urgent_reminder_hours"] = int(saved.get("urgent_reminder_hours", 24))
    st.session_state["model"] = saved.get("model", "")
    st.session_state["graph_token_input"] = saved.get("graph_token", "")
    st.session_state["days"] = int(saved.get("days", 10))
    st.session_state["top"] = int(saved.get("top", 200))
    st.session_state["custom_prompt"] = saved.get("custom_prompt", "")
    st.session_state["priority_senders_raw"] = saved.get("priority_senders_raw", "")
    st.session_state["calendar_timezone"] = saved.get("calendar_timezone", "Europe/Prague")
    st.session_state["auto_save_settings"] = bool(saved.get("auto_save_settings", True))
    st.session_state["settings_initialized"] = True


def build_client(api_key: str, base_url: str, timeout_seconds: int = 60) -> OpenAI:
    # Keep retries low so blocked corporate networks fail fast with a clear message.
    return OpenAI(api_key=api_key, base_url=base_url, timeout=timeout_seconds, max_retries=1)


def merge_prompt(custom_prompt: str, senders: list[str]) -> str:
    extra = []
    if senders:
        extra.append("Preferovaní odesílatelé: " + ", ".join(senders))
    if extra:
        return custom_prompt + "\n\n" + "\n".join(extra)
    return custom_prompt


def parse_json_content(content: str) -> dict:
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        if "```" in content:
            cleaned = content.replace("```json", "").replace("```", "").strip()
            return json.loads(cleaned)
        raise


def list_models(client: OpenAI) -> list[str]:
    resp = client.models.list()
    return sorted([m.id for m in getattr(resp, "data", []) if getattr(m, "id", None)])


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
                "subject": m.get("subject", "(bez předmětu)"),
                "from": sender,
                "receivedDateTime": m.get("receivedDateTime", ""),
                "bodyPreview": m.get("bodyPreview", ""),
                "importance": m.get("importance", "normal"),
                "categories": m.get("categories", []),
            }
        )
    return items


def _parse_graph_datetime(value: str) -> datetime:
    try:
        return datetime.fromisoformat((value or "").replace("Z", "+00:00"))
    except Exception:
        return datetime.now(timezone.utc)


def filter_items_for_analysis(items: list[dict], mode: str, urgent_reminder_hours: int) -> tuple[list[dict], dict]:
    if mode == "Všechny (včetně již označených)":
        return items, {"skipped_labeled": 0, "urgent_reincluded": 0}

    unlabeled = []
    labeled = []
    for item in items:
        categories = item.get("categories") or []
        if any(str(cat).startswith(MAILAI_CATEGORY_PREFIX) for cat in categories):
            labeled.append(item)
        else:
            unlabeled.append(item)

    if mode == "Jen bez MailAI štítku":
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


def fetch_unread_messages(token: str, days: int, top: int) -> list[dict]:
    since = (datetime.now(timezone.utc) - timedelta(days=days)).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    messages = _fetch_graph_messages(
        token,
        "https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages",
        {
            "$select": "id,conversationId,subject,from,receivedDateTime,bodyPreview,isRead,importance,categories",
            "$top": top,
            "$orderby": "receivedDateTime DESC",
            "$filter": f"isRead eq false and receivedDateTime ge {since}",
        },
        top,
    )
    return _normalize_inbox_items(messages)


def fetch_not_replied_messages(token: str, days: int, top: int) -> list[dict]:
    since = (datetime.now(timezone.utc) - timedelta(days=days)).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    inbox_messages = _fetch_graph_messages(
        token,
        "https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages",
        {
            "$select": "id,conversationId,subject,from,receivedDateTime,bodyPreview,isRead,importance,categories",
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


def summarize_unread(client: OpenAI, model: str, prompt: str, items: list[dict], days: int) -> dict:
    prepared_items = []
    for itm in items:
        prepared_items.append(
            {
                "id": str(itm.get("id") or ""),
                "conversationId": str(itm.get("conversationId") or ""),
                "subject": str(itm.get("subject") or "(bez předmětu)"),
                "from": str(itm.get("from") or "(neznámý odesílatel)"),
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
        # Fallback for models/providers that reject response_format=json_object.
        resp = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": prompt + "\n\nVrátíš pouze validní JSON bez markdownu."},
                {"role": "user", "content": user_text},
            ],
            temperature=0.2,
            max_tokens=3500,
        )
        return parse_json_content(resp.choices[0].message.content)


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


def graph_patch_read(token: str, msg_id: str) -> None:
    r = requests.patch(
        f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
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
) -> dict:
    payload = {
        "subject": subject,
        "body": {
            "contentType": "Text",
            "content": body_text,
        },
        "start": {
            "dateTime": start_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": timezone_name,
        },
        "end": {
            "dateTime": end_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": timezone_name,
        },
        "isReminderOn": True,
        "categories": [MAILAI_DEADLINE_CATEGORY[0]],
    }
    r = requests.post(
        "https://graph.microsoft.com/v1.0/me/events",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=20,
    )
    r.raise_for_status()
    return r.json()


def graph_get_master_categories(token: str) -> set[str]:
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
        headers={"Authorization": f"Bearer {token}"},
        timeout=20,
    )
    r.raise_for_status()
    return {c.get("displayName", "") for c in r.json().get("value", []) if c.get("displayName")}


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


def graph_create_master_category(token: str, name: str, color: str) -> None:
    r = requests.post(
        "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={
            "displayName": name,
            "color": color,
        },
        timeout=20,
    )
    if r.status_code not in (200, 201, 409):
        r.raise_for_status()


def ensure_mailai_master_categories(token: str) -> None:
    existing = graph_get_master_categories(token)
    all_categories = list(MAILAI_CATEGORY_MAP.values()) + [MAILAI_DEADLINE_CATEGORY]
    for name, color in all_categories:
        if name not in existing:
            graph_create_master_category(token, name, color)


def is_master_categories_forbidden(error: Exception) -> bool:
    if not isinstance(error, requests.HTTPError):
        return False
    response = error.response
    if not response:
        return False
    req_url = getattr(response.request, "url", "") or ""
    return response.status_code == 403 and "masterCategories" in req_url


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
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json={"categories": current_categories + [category_name]},
        timeout=20,
    )
    patch_r.raise_for_status()


def initialize_bucket_overrides(result: dict) -> None:
    new_ids = set()
    for bucket_key in BUCKET_ORDER:
        for itm in (result.get("buckets") or {}).get(bucket_key, []):
            msg_id = str(itm.get("id") or "")
            if not msg_id:
                continue
            new_ids.add(msg_id)
            st.session_state[f"bucket_override_{msg_id}"] = bucket_key

    for old_id in st.session_state.get("editable_bucket_ids", []):
        if old_id not in new_ids:
            st.session_state.pop(f"bucket_override_{old_id}", None)

    st.session_state["editable_bucket_ids"] = sorted(new_ids)


def get_bucket_overrides(result: dict) -> dict[str, str]:
    editable_ids = st.session_state.get("editable_bucket_ids") or []
    if not editable_ids:
        initialize_bucket_overrides(result)
        editable_ids = st.session_state.get("editable_bucket_ids") or []

    overrides = {}
    for msg_id in editable_ids:
        bucket_key = st.session_state.get(f"bucket_override_{msg_id}")
        if bucket_key in BUCKET_ORDER:
            overrides[msg_id] = bucket_key
    return overrides


def build_effective_buckets(result: dict, overrides: dict[str, str]) -> dict[str, list[dict]]:
    effective = {bucket_key: [] for bucket_key in BUCKET_ORDER}
    seen_ids = set()
    for source_bucket in BUCKET_ORDER:
        for itm in (result.get("buckets") or {}).get(source_bucket, []):
            msg_id = str(itm.get("id") or "")
            if msg_id:
                if msg_id in seen_ids:
                    continue
                seen_ids.add(msg_id)
            target_bucket = overrides.get(msg_id, source_bucket)
            if target_bucket not in effective:
                target_bucket = source_bucket

            cloned = dict(itm)
            cloned["suggested_bucket"] = source_bucket
            cloned["selected_bucket"] = target_bucket
            effective[target_bucket].append(cloned)
    return effective


def render_bucket(bucket_key: str, items: list[dict], editable: bool = False):
    cfg = BUCKET_UI.get(bucket_key, {"label": bucket_key, "color": "#888", "emoji": ""})
    label = cfg["label"]
    color = cfg["color"]
    emoji = cfg["emoji"]
    count = len(items)
    st.markdown(
        f'<div style="border-left: 5px solid {color}; padding: 4px 12px; margin-bottom: 4px;">'
        f'<span style="color:{color}; font-size: 1.1rem; font-weight: 700">{emoji} {label} ({count})</span>'
        f'</div>',
        unsafe_allow_html=True,
    )
    if not items:
        st.caption("Bez položek")
        return
    for itm in items[:50]:
        msg_id = str(itm.get("id") or "")
        deadline_badge = ""
        if itm.get("has_deadline"):
            hint = itm.get("deadline_hint") or "termín"
            deadline_badge = f' <span style="background:#7d3c98;color:#fff;border-radius:4px;padding:1px 6px;font-size:0.8rem">📅 {hint}</span>'
        moved_badge = ""
        if itm.get("suggested_bucket") and itm.get("suggested_bucket") != bucket_key:
            original_label = BUCKET_UI.get(itm["suggested_bucket"], {}).get("label", itm["suggested_bucket"])
            moved_badge = (
                ' <span style="background:#ecf0f1;color:#2c3e50;border-radius:4px;padding:1px 6px;font-size:0.8rem">'
                f'Původně: {original_label}</span>'
            )

        if editable and msg_id:
            col_info, col_choice = st.columns([5, 2])
            with col_info:
                st.markdown(
                    f'<span style="color:{color}">●</span> **{itm.get("subject", "")}** | '
                    f'<span style="color:#888">{itm.get("from", "")}</span>{deadline_badge}{moved_badge}',
                    unsafe_allow_html=True,
                )
                if itm.get("reason"):
                    st.caption(itm["reason"])
            with col_choice:
                st.selectbox(
                    "Cílová kategorie",
                    options=list(BUCKET_ORDER),
                    format_func=lambda value: BUCKET_UI[value]["label"],
                    key=f"bucket_override_{msg_id}",
                    label_visibility="collapsed",
                )
        else:
            st.markdown(
                f'<span style="color:{color}">●</span> **{itm.get("subject", "")}** | '
                f'<span style="color:#888">{itm.get("from", "")}</span>{deadline_badge}{moved_badge}',
                unsafe_allow_html=True,
            )
            if itm.get("reason"):
                st.caption(itm["reason"])


def get_deadline_items(effective_buckets: dict[str, list[dict]]) -> list[dict]:
    deadline_items = []
    seen_ids = set()
    for bucket_key in BUCKET_ORDER:
        for itm in effective_buckets.get(bucket_key, []):
            if not itm.get("has_deadline"):
                continue
            msg_id = str(itm.get("id") or "")
            if not msg_id or msg_id in seen_ids:
                continue
            seen_ids.add(msg_id)
            cloned = dict(itm)
            cloned["bucket_key"] = bucket_key
            deadline_items.append(cloned)
    return deadline_items


def parse_deadline_date_hint(hint: str | None) -> datetime | None:
    text = (hint or "").strip().lower()
    if not text:
        return None

    now_local = datetime.now().astimezone()

    # Relative Czech hints.
    if "dnes" in text:
        return now_local
    if "zítra" in text or "zitra" in text:
        return now_local + timedelta(days=1)
    if "pozítří" in text or "pozitri" in text:
        return now_local + timedelta(days=2)

    # Common Czech date formats: 12. 4. 2026, 12.04.2026, 12/04/2026, 12-04-2026
    m = re.search(r"\b(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{2,4})\b", text)
    if m:
        day = int(m.group(1))
        month = int(m.group(2))
        year = int(m.group(3))
        if year < 100:
            year += 2000
        try:
            return now_local.replace(year=year, month=month, day=day)
        except ValueError:
            pass

    # Day+month without year (assume current year, otherwise next year if already passed).
    m = re.search(r"\b(\d{1,2})\s*[./-]\s*(\d{1,2})(?:\s|$)", text)
    if m:
        day = int(m.group(1))
        month = int(m.group(2))
        year = now_local.year
        try:
            candidate = now_local.replace(year=year, month=month, day=day)
            if candidate.date() < now_local.date():
                candidate = candidate.replace(year=year + 1)
            return candidate
        except ValueError:
            pass

    return None


def main():
    st.set_page_config(page_title="MailAI Local", layout="wide")
    st.title("MailAI Local")
    st.caption("Lokální alternativa bez Outlook add-inu")
    initialize_state_from_settings()

    with st.sidebar:
        st.header("Nastavení")
        llm_api_key = st.text_input("LLM API key", type="password", key="llm_api_key")
        llm_base_url = st.text_input("LLM Base URL", key="llm_base_url")
        llm_timeout = st.number_input("LLM timeout (sekundy)", min_value=10, max_value=600, key="llm_timeout")
        analysis_mode = st.selectbox(
            "Režim výběru e-mailů",
            options=["Nepřečtené", "Bez odpovědi (Inbox vs Sent)"],
            key="analysis_mode",
        )
        label_handling_mode = st.selectbox(
            "Práce s již označenými e-maily",
            options=[
                "Jen bez MailAI štítku + urgentní připomenutí",
                "Jen bez MailAI štítku",
                "Všechny (včetně již označených)",
            ],
            key="label_handling_mode",
        )
        urgent_reminder_hours = st.number_input(
            "Připomenout urgentní po (hodin)",
            min_value=1,
            max_value=240,
            key="urgent_reminder_hours",
        )
        graph_token = st.text_input("Graph Access Token", type="password", key="graph_token_input")
        calendar_timezone = st.text_input("Časová zóna kalendáře", key="calendar_timezone")
        days = st.number_input("Počet dní zpět", min_value=1, max_value=30, key="days")
        top = st.number_input("Max počet e-mailů", min_value=10, max_value=1000, key="top")
        custom_prompt = st.text_area("Custom prompt", height=120, key="custom_prompt")
        priority_senders_raw = st.text_area("Preferovaní odesílatelé", height=80, key="priority_senders_raw")
        auto_save_settings = st.checkbox("Automaticky ukládat nastavení lokálně", key="auto_save_settings")

        col_save, col_clear = st.columns(2)
        if col_save.button("Uložit"):
            save_local_settings(build_settings_payload())
            st.success("Nastavení uloženo do .mailai_local_settings.json")
        if col_clear.button("Smazat"):
            if SETTINGS_FILE.exists():
                SETTINGS_FILE.unlink()
            st.success("Lokální uložené nastavení smazáno")

        if st.button("Načíst modely"):
            try:
                client = build_client(llm_api_key, llm_base_url, int(llm_timeout))
                models = list_models(client)
                st.session_state["models"] = models
                st.success(f"Načteno modelů: {len(models)}")
            except Exception as e:
                st.error(f"Chyba načítání modelů: {e}")

        if st.button("Ověřit Graph oprávnění"):
            if not graph_token:
                st.error("Nejdřív vlož Graph Access Token")
            else:
                claims = decode_jwt_claims_unverified(graph_token)
                scp = claims.get("scp", "")
                roles = claims.get("roles", [])
                aud = claims.get("aud", "")

                st.markdown("### Diagnostika Graph tokenu")
                st.write(f"aud: {aud}")
                st.write(f"scp: {scp or '(není v tokenu)'}")
                if roles:
                    st.write(f"roles: {roles}")

                me_status, me_msg = graph_endpoint_status(
                    graph_token,
                    "https://graph.microsoft.com/v1.0/me",
                    {"$select": "id,userPrincipalName"},
                )
                msg_status, msg_msg = graph_endpoint_status(
                    graph_token,
                    "https://graph.microsoft.com/v1.0/me/messages",
                    {"$top": 1, "$select": "id"},
                )
                cat_status, cat_msg = graph_endpoint_status(
                    graph_token,
                    "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
                )
                evt_status, evt_msg = graph_endpoint_status(
                    graph_token,
                    "https://graph.microsoft.com/v1.0/me/events",
                    {"$top": 1, "$select": "id"},
                )

                st.write(f"/me: {me_status} - {me_msg}")
                st.write(f"/me/messages: {msg_status} - {msg_msg}")
                st.write(f"/me/outlook/masterCategories: {cat_status} - {cat_msg}")
                st.write(f"/me/events: {evt_status} - {evt_msg}")
                st.caption(
                    "Pro masterCategories je potřeba MailboxSettings.ReadWrite a pro kalendář Calendars.ReadWrite. "
                    "Po změně oprávnění vždy vygeneruj nový token."
                )

        models = st.session_state.get("models", [])
        current_model = st.session_state.get("model", "")
        if models:
            model_options = ["(Vlastní model)"] + models
            default_choice = current_model if current_model in models else "(Vlastní model)"
            selected_choice = st.selectbox(
                "Model",
                options=model_options,
                index=model_options.index(default_choice),
                key="model_picker",
            )
            if selected_choice == "(Vlastní model)":
                custom_model_value = current_model if current_model not in models else ""
                chosen_model = st.text_input("Vlastní název modelu", value=custom_model_value, key="model_custom")
            else:
                chosen_model = selected_choice
        else:
            chosen_model = st.text_input("Model", value=current_model, key="model_custom_no_list")

        if chosen_model != current_model:
            st.session_state["model"] = chosen_model

    final_model = st.session_state.get("model", "")

    st.markdown("### Inbox souhrn (nepřečtené e-maily)")
    if st.button("Analyzovat nepřečtené e-maily za posledních N dní", type="primary"):
        if not llm_api_key:
            st.error("Zadej LLM API key")
            return
        if not graph_token:
            st.error("Zadej Graph Access Token")
            return
        if not final_model:
            st.error("Zadej nebo vyber model")
            return

        senders = [s.strip() for s in priority_senders_raw.replace(",", "\n").split("\n") if s.strip()]
        prompt = merge_prompt(custom_prompt or INBOX_PROMPT, senders)

        try:
            if auto_save_settings:
                save_local_settings(build_settings_payload())

            with st.spinner("Načítám e-maily z Microsoft Graph..."):
                if analysis_mode == "Bez odpovědi (Inbox vs Sent)":
                    items = fetch_not_replied_messages(graph_token, int(days), int(top))
                    st.info(f"Načteno e-mailů bez odpovědi: {len(items)}")
                else:
                    items = fetch_unread_messages(graph_token, int(days), int(top))
                    st.info(f"Načteno nepřečtených e-mailů: {len(items)}")

            items, filter_stats = filter_items_for_analysis(
                items,
                label_handling_mode,
                int(urgent_reminder_hours),
            )
            if label_handling_mode != "Všechny (včetně již označených)":
                st.info(
                    f"Po filtraci štítků do analýzy: {len(items)} | přeskočeno již označených: {filter_stats['skipped_labeled']}"
                )
                if filter_stats["urgent_reincluded"]:
                    st.info(
                        f"Urgentní připomenutí vráceno do analýzy: {filter_stats['urgent_reincluded']}"
                    )

            if not items:
                st.warning("Po filtraci nezbyly žádné e-maily pro analýzu.")
                return

            llm_items = items[:MAX_EMAILS_FOR_LLM]
            if len(items) > MAX_EMAILS_FOR_LLM:
                st.warning(
                    f"Pro LLM analýzu používám prvních {MAX_EMAILS_FOR_LLM} e-mailů z {len(items)} kvůli rychlosti a stabilitě."
                )

            with st.spinner("Analyzuji přes LLM..."):
                client = build_client(llm_api_key, llm_base_url, int(llm_timeout))
                result = summarize_unread(client, final_model, prompt, llm_items, int(days))
                result = enforce_no_delete_policy(result)

            st.session_state["inbox_result"] = result
            st.session_state["graph_token"] = graph_token
            initialize_bucket_overrides(result)
            st.success("Souhrn hotový")
        except APITimeoutError:
            st.error(
                "LLM timeout: provider neodpověděl včas. Zkus jiný model, ověř Base URL, nebo zvyšit timeout v nastavení."
            )
        except Exception as e:
            msg = str(e)
            if "required attributes" in msg.lower() or "požadovaných atribut" in msg.lower():
                st.error(
                    "Model odmítl formát vstupu/výstupu. Zkus chat model (např. gpt-4o-mini) a klikni znovu na Načíst modely."
                )
                st.caption(f"Detail chyby: {msg}")
            else:
                st.error(f"Chyba: {e}")

    result = st.session_state.get("inbox_result")
    if result:
        bucket_overrides = get_bucket_overrides(result)
        effective_buckets = build_effective_buckets(result, bucket_overrides)
        effective_counts = {bucket_key: len(effective_buckets[bucket_key]) for bucket_key in BUCKET_ORDER}

        st.markdown("### Výsledek")
        st.write(result.get("overview", ""))

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Urgentní", effective_counts.get("urgentni", 0))
        c2.metric("Středně důležité", effective_counts.get("stredne_dulezite", 0))
        c3.metric("Počká", effective_counts.get("pocka", 0))
        c4.metric("K přeposlání", effective_counts.get("k_preposlani", 0))
        c5.metric("Ignorovat", effective_counts.get("ignorovat", 0))

        st.markdown("### Aktuální rozdělení (možno upravit)")
        st.caption("Kategorie můžeš měnit přímo u každého e-mailu v tomto přehledu.")
        for bkey in BUCKET_ORDER:
            render_bucket(bkey, effective_buckets.get(bkey, []), editable=True)

        deadline_items = get_deadline_items(effective_buckets)
        if deadline_items:
            st.markdown("### Termíny do kalendáře")
            st.caption("U e-mailů s termínem můžeš jedním klikem vytvořit událost v Outlook kalendáři.")
            for itm in deadline_items:
                msg_id = str(itm.get("id") or "")
                start_default = _parse_graph_datetime(itm.get("receivedDateTime", ""))
                if start_default.tzinfo:
                    start_default = start_default.astimezone().replace(tzinfo=None)
                start_default = start_default.replace(second=0, microsecond=0) + timedelta(hours=1)
                hint_date = parse_deadline_date_hint(itm.get("deadline_hint"))
                if hint_date is not None:
                    start_default = start_default.replace(
                        year=hint_date.year,
                        month=hint_date.month,
                        day=hint_date.day,
                    )

                date_key = f"event_date_{msg_id}"
                time_key = f"event_time_{msg_id}"
                dur_key = f"event_dur_{msg_id}"

                col_info, col_date, col_time, col_dur, col_btn = st.columns([4, 1.3, 1.2, 1, 1.4])
                with col_info:
                    hint = itm.get("deadline_hint") or "bez upřesnění"
                    st.markdown(
                        f"**{itm.get('subject', '(bez předmětu)')}**  \n"
                        f"<span style='color:#666'>{itm.get('from', '(neznámý odesílatel)')}</span>"
                        f"<br><span style='color:#7d3c98'>📅 {hint}</span>",
                        unsafe_allow_html=True,
                    )
                with col_date:
                    st.date_input(
                        "Datum",
                        value=start_default.date(),
                        key=date_key,
                        format="DD/MM/YYYY",
                        label_visibility="collapsed",
                    )
                with col_time:
                    st.time_input("Čas", value=start_default.time(), key=time_key, label_visibility="collapsed")
                with col_dur:
                    st.number_input(
                        "Min",
                        min_value=15,
                        max_value=480,
                        step=15,
                        value=30,
                        key=dur_key,
                        label_visibility="collapsed",
                    )
                with col_btn:
                    if st.button("Vložit do kalendáře", key=f"event_btn_{msg_id}"):
                        start_dt = datetime.combine(st.session_state[date_key], st.session_state[time_key])
                        end_dt = start_dt + timedelta(minutes=int(st.session_state[dur_key]))
                        body_text = (
                            f"MailAI termín z e-mailu\n"
                            f"Od: {itm.get('from', '')}\n"
                            f"Předmět: {itm.get('subject', '')}\n"
                            f"Deadline hint: {itm.get('deadline_hint') or ''}\n"
                            f"Důvod: {itm.get('reason') or ''}\n"
                            f"Message ID: {msg_id}"
                        )
                        try:
                            graph_create_calendar_event(
                                token,
                                f"Termín: {itm.get('subject', '(bez předmětu)')}",
                                start_dt,
                                end_dt,
                                calendar_timezone,
                                body_text,
                            )
                            st.success(f"Událost vytvořena pro: {itm.get('subject', '(bez předmětu)')}")
                        except Exception as e:
                            st.error(f"Nepodařilo se vytvořit událost v kalendáři: {e}")

        st.markdown("### Doporučené hromadné akce")
        actions = result.get("recommended_bulk_actions", {})
        mark_ids = actions.get("mark_read_ids", [])

        st.write(f"Označit jako přečtené: {len(mark_ids)}")
        st.write("Smazat: 0 (zakázáno politikou aplikace)")

        token = st.session_state.get("graph_token", "")
        if token:
            if st.button("Přiřadit štítky podle aktuálního rozdělení"):
                ok = 0
                fail = 0
                categories_prepared = False
                try:
                    ensure_mailai_master_categories(token)
                    categories_prepared = True
                except Exception as e:
                    if is_master_categories_forbidden(e):
                        st.warning(
                            "Graph token nemá oprávnění pro správu kategorií (masterCategories). "
                            "Přidej scope MailboxSettings.ReadWrite a vygeneruj nový token. "
                            "Pokusím se pokračovat: pokud kategorie už existují, přiřazení může fungovat."
                        )
                    else:
                        st.error(f"Nepodařilo se připravit Outlook kategorie: {e}")
                        st.stop()

                deadline_cat_name = MAILAI_DEADLINE_CATEGORY[0]
                for bucket_key, (category_name, _) in MAILAI_CATEGORY_MAP.items():
                    for itm in effective_buckets.get(bucket_key, []):
                        msg_id = itm.get("id")
                        if not msg_id:
                            continue
                        try:
                            graph_assign_category(token, msg_id, category_name)
                            ok += 1
                        except Exception:
                            fail += 1
                        # Přiřaď navíc štítek S termínem urgentním a středně důležitým s termínem
                        if bucket_key in ("urgentni", "stredne_dulezite") and itm.get("has_deadline"):
                            try:
                                graph_assign_category(token, msg_id, deadline_cat_name)
                            except Exception:
                                pass

                st.success(f"Štítek přiřazen u {ok} e-mailů")
                if fail:
                    st.warning(f"Nepodařilo se přiřadit štítek u {fail} e-mailů")
                if not categories_prepared and ok == 0:
                    st.info(
                        "Kategorie pravděpodobně v mailboxu neexistují. Po přidání oprávnění "
                        "MailboxSettings.ReadWrite je aplikace vytvoří automaticky."
                    )

            if st.button("Provést doporučené označení jako přečtené"):
                ok = 0
                fail = 0
                for msg_id in mark_ids:
                    try:
                        graph_patch_read(token, msg_id)
                        ok += 1
                    except Exception:
                        fail += 1
                st.success(f"Označeno jako přečtené: {ok}")
                if fail:
                    st.warning(f"Nepovedlo se označit: {fail}")

            st.caption(
                "Pro hromadné akce je potřeba Mail.ReadWrite. Pro vytváření Outlook kategorií je potřeba "
                "MailboxSettings.ReadWrite. Pro vložení termínu do kalendáře je potřeba Calendars.ReadWrite."
            )


if __name__ == "__main__":
    main()

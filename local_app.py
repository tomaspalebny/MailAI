import json
from datetime import datetime, timedelta, timezone

import requests
import streamlit as st
from openai import OpenAI


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
    "urgentni": [{"id":"...","subject":"...","from":"...","reason":"...","action":"..."}],
    "stredne_dulezite": [{"id":"...","subject":"...","from":"...","reason":"...","action":"..."}],
    "pocka": [{"id":"...","subject":"...","from":"...","reason":"...","action":"..."}],
    "k_preposlani": [{"id":"...","subject":"...","from":"...","reason":"...","forward_to":"...","action":"..."}],
    "ignorovat": [{"id":"...","subject":"...","from":"...","reason":"...","action":"smazat|oznacit_jako_prectene|ignorovat"}]
  },
  "recommended_bulk_actions": {
    "mark_read_ids": ["..."],
    "delete_ids": ["..."]
  }
}
Pravidla:
- Použij jen kategorie: urgentni, stredne_dulezite, pocka, k_preposlani, ignorovat.
- Každý e-mail zařaď právě do jedné kategorie.
- Buď konzervativní u mazání: do delete dávej jen zjevný spam/newsletter bez akční hodnoty.
- Odpovídej česky.
"""


def build_client(api_key: str, base_url: str) -> OpenAI:
    return OpenAI(api_key=api_key, base_url=base_url)


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


def fetch_unread_messages(token: str, days: int, top: int) -> list[dict]:
    since = (datetime.now(timezone.utc) - timedelta(days=days)).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    r = requests.get(
        "https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages",
        headers={"Authorization": f"Bearer {token}"},
        params={
            "$select": "id,subject,from,receivedDateTime,bodyPreview,isRead,importance",
            "$top": top,
            "$orderby": "receivedDateTime DESC",
            "$filter": f"isRead eq false and receivedDateTime ge {since}",
        },
        timeout=30,
    )
    r.raise_for_status()
    msgs = r.json().get("value", [])
    items = []
    for m in msgs:
        ea = m.get("from", {}).get("emailAddress", {})
        sender = f"{ea.get('name', '')} <{ea.get('address', '')}>".strip()
        items.append(
            {
                "id": m.get("id"),
                "subject": m.get("subject", "(bez předmětu)"),
                "from": sender,
                "receivedDateTime": m.get("receivedDateTime", ""),
                "bodyPreview": m.get("bodyPreview", ""),
                "importance": m.get("importance", "normal"),
            }
        )
    return items


def summarize_unread(client: OpenAI, model: str, prompt: str, items: list[dict], days: int) -> dict:
    user_text = json.dumps(
        {
            "window_days": days,
            "total_unread_fetched": len(items),
            "emails": items,
        },
        ensure_ascii=False,
    )
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


def graph_delete(token: str, msg_id: str) -> None:
    r = requests.delete(
        f"https://graph.microsoft.com/v1.0/me/messages/{msg_id}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=20,
    )
    r.raise_for_status()


def render_bucket(title: str, items: list[dict]):
    st.subheader(f"{title} ({len(items)})")
    if not items:
        st.caption("Bez položek")
        return
    for itm in items[:50]:
        st.markdown(f"- **{itm.get('subject', '')}** | {itm.get('from', '')}")
        if itm.get("reason"):
            st.caption(itm["reason"])


def main():
    st.set_page_config(page_title="MailAI Local", layout="wide")
    st.title("MailAI Local")
    st.caption("Lokální alternativa bez Outlook add-inu")

    with st.sidebar:
        st.header("Nastavení")
        llm_api_key = st.text_input("LLM API key", type="password")
        llm_base_url = st.text_input("LLM Base URL", value="https://llm.ai.e-infra.cz/v1/")
        model = st.text_input("Model", value="")
        graph_token = st.text_input("Graph Access Token", type="password")
        days = st.number_input("Počet dní zpět", min_value=1, max_value=30, value=10)
        top = st.number_input("Max počet e-mailů", min_value=10, max_value=500, value=200)
        custom_prompt = st.text_area("Custom prompt", value="", height=120)
        priority_senders_raw = st.text_area("Preferovaní odesílatelé", value="", height=80)

        if st.button("Načíst modely"):
            try:
                client = build_client(llm_api_key, llm_base_url)
                models = list_models(client)
                st.session_state["models"] = models
                st.success(f"Načteno modelů: {len(models)}")
            except Exception as e:
                st.error(f"Chyba načítání modelů: {e}")

        models = st.session_state.get("models", [])
        if models:
            selected = st.selectbox("Vyber model z načtených", options=models, index=0)
            if st.button("Použít vybraný model"):
                st.session_state["selected_model"] = selected
                st.success(f"Model nastaven: {selected}")

    final_model = st.session_state.get("selected_model") or model

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
            with st.spinner("Načítám e-maily z Microsoft Graph..."):
                items = fetch_unread_messages(graph_token, int(days), int(top))
            st.info(f"Načteno nepřečtených e-mailů: {len(items)}")

            with st.spinner("Analyzuji přes LLM..."):
                client = build_client(llm_api_key, llm_base_url)
                result = summarize_unread(client, final_model, prompt, items, int(days))

            st.session_state["inbox_result"] = result
            st.session_state["graph_token"] = graph_token
            st.success("Souhrn hotový")
        except Exception as e:
            st.error(f"Chyba: {e}")

    result = st.session_state.get("inbox_result")
    if result:
        st.markdown("### Výsledek")
        st.write(result.get("overview", ""))

        counts = result.get("counts", {})
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Urgentní", counts.get("urgentni", 0))
        c2.metric("Středně důležité", counts.get("stredne_dulezite", 0))
        c3.metric("Počká", counts.get("pocka", 0))
        c4.metric("K přeposlání", counts.get("k_preposlani", 0))
        c5.metric("Ignorovat", counts.get("ignorovat", 0))

        buckets = result.get("buckets", {})
        render_bucket("Urgentní", buckets.get("urgentni", []))
        render_bucket("Středně důležité", buckets.get("stredne_dulezite", []))
        render_bucket("Počká", buckets.get("pocka", []))
        render_bucket("K přeposlání", buckets.get("k_preposlani", []))
        render_bucket("Ignorovat", buckets.get("ignorovat", []))

        st.markdown("### Doporučené hromadné akce")
        actions = result.get("recommended_bulk_actions", {})
        mark_ids = actions.get("mark_read_ids", [])
        del_ids = actions.get("delete_ids", [])

        st.write(f"Označit jako přečtené: {len(mark_ids)}")
        st.write(f"Smazat: {len(del_ids)}")

        token = st.session_state.get("graph_token", "")
        if token:
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

            if st.button("Provést doporučené smazání"):
                ok = 0
                fail = 0
                for msg_id in del_ids:
                    try:
                        graph_delete(token, msg_id)
                        ok += 1
                    except Exception:
                        fail += 1
                st.success(f"Smazáno: {ok}")
                if fail:
                    st.warning(f"Nepovedlo se smazat: {fail}")

            st.caption("Pro hromadné akce je obvykle potřeba Graph oprávnění Mail.ReadWrite.")


if __name__ == "__main__":
    main()

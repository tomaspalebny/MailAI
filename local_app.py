import json
import base64
import re
import html
import os
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlparse, urlunparse

import requests
import streamlit as st
try:
    import msal
except ImportError:
    msal = None
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

INBOX_PROMPT_EN = """You are an email assistant. Process the list of unread emails and return only valid JSON with this structure:
{
  "overview": "brief summary",
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
    "ignorovat": [{"id":"...","subject":"...","from":"...","reason":"...","action":"mark_as_read|ignore"}]
  },
  "recommended_bulk_actions": {
    "mark_read_ids": ["..."]
  }
}
Rules:
- Use only these categories: urgentni, stredne_dulezite, pocka, k_preposlani, ignorovat.
- Place each email in exactly one category.
- Never suggest deleting emails or a delete action.
- For urgentni and stredne_dulezite: has_deadline=true if the email contains a specific deadline/date (deadline, closing, meeting, by when). deadline_hint = brief description of the deadline in English, or null.
- Respond in English.
"""

MAX_EMAILS_FOR_LLM = 120
SETTINGS_FILE = Path(".mailai_local_settings.json")
MSAL_CACHE_FILE = Path(".mailai_msal_cache.json")
MSAL_SCOPES = [
    "Mail.ReadWrite",
    "MailboxSettings.ReadWrite",
    "Calendars.ReadWrite",
    "offline_access",
]
DEFAULT_MS_TENANT_ID = os.getenv("MAILAI_MS_TENANT_ID", "common")
DEFAULT_MS_CLIENT_ID = os.getenv("MAILAI_MS_CLIENT_ID", "")
MAILAI_CATEGORY_PREFIX = "MailAI/"

# Internal keys for analysis / label-handling modes (stored in settings as-is)
ANALYSIS_MODE_UNREAD = "Nepřečtené"
ANALYSIS_MODE_NOT_REPLIED = "Bez odpovědi (Inbox vs Sent)"
LABEL_MODE_DEFAULT = "Jen bez MailAI štítku + urgentní připomenutí"
LABEL_MODE_WITHOUT = "Jen bez MailAI štítku"
LABEL_MODE_ALL = "Všechny (včetně již označených)"
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

# ---------------------------------------------------------------------------
# Internationalisation
# ---------------------------------------------------------------------------

TRANSLATIONS: dict[str, dict[str, str]] = {
    "cs": {
        # Language selector
        "language_label": "Jazyk / Language",
        # Page
        "page_subtitle": "Lokální alternativa bez Outlook add-inu",
        # Sidebar
        "sidebar_header": "Nastavení",
        "system_prompt_expander": "Zobrazit systémový prompt",
        "system_prompt_label": "Systémový prompt",
        "system_prompt_help": "Hlavní instrukce pro AI analýzu. Ukládá se do lokálního nastavení.",
        "reset_prompt_btn": "Obnovit výchozí systémový prompt",
        "reset_prompt_help": "Vrátí systémový prompt na původní výchozí hodnotu.",
        "llm_timeout_label": "LLM timeout (sekundy)",
        "analysis_mode_label": "Režim výběru e-mailů",
        "analysis_mode_unread": "Nepřečtené",
        "analysis_mode_not_replied": "Bez odpovědi (Inbox vs Sent)",
        "label_mode_label": "Práce s již označenými e-maily",
        "label_mode_default": "Jen bez MailAI štítku + urgentní připomenutí",
        "label_mode_without": "Jen bez MailAI štítku",
        "label_mode_all": "Všechny (včetně již označených)",
        "urgent_reminder_label": "Připomenout urgentní po (hodin)",
        "calendar_tz_label": "Časová zóna kalendáře",
        "days_label": "Počet dní zpět",
        "top_label": "Max počet e-mailů",
        "custom_prompt_label": "Custom prompt",
        "custom_prompt_help": "Přidá se k výchozím instrukcím pro analýzu, nenahrazuje je.",
        "priority_senders_label": "Preferovaní odesílatelé",
        "auto_save_label": "Automaticky ukládat nastavení lokálně",
        "prompt_validation_warning": "Validace systémového promptu našla problém, který může rozbít analýzu:",
        "prompt_validation_ok": "Systémový prompt prošel základní validací.",
        "prompt_preview_expander": "Náhled finálního promptu pro model",
        "save_btn": "Uložit",
        "clear_btn": "Smazat",
        "settings_saved": "Nastavení uloženo do .mailai_local_settings.json",
        "settings_cleared": "Lokální uložené nastavení smazáno",
        "load_models_btn": "Načíst modely",
        "models_loaded": "Načteno modelů: {n}",
        "models_load_error": "Chyba načítání modelů: {e}",
        "verify_graph_btn": "Ověřit Graph oprávnění",
        "graph_token_label": "Graph Access Token (ručně, volitelné)",
        "no_graph_token_error": "Nejdřív vlož Graph Access Token nebo se přihlas přes OAuth",
        "oauth_header": "Microsoft OAuth (delší platnost tokenu)",
        "oauth_client_id_label": "Azure Client ID",
        "oauth_tenant_id_label": "Azure Tenant ID",
        "oauth_signin_btn": "Přihlásit přes Microsoft",
        "oauth_signout_btn": "Odhlásit Microsoft účet",
        "oauth_logged_in": "OAuth přihlášen: {account}",
        "oauth_logged_out": "OAuth nepřihlášen",
        "oauth_missing_client_id": "Pro OAuth vyplň Azure Client ID.",
        "oauth_login_failed": "OAuth přihlášení selhalo: {e}",
        "oauth_cache_cleared": "OAuth cache byla smazána.",
        "oauth_source_manual": "ruční token",
        "oauth_source_oauth": "OAuth (MSAL cache)",
        "oauth_source_none": "není dostupný",
        "oauth_token_source": "Zdroj Graph tokenu: {source}",
        "oauth_scopes_caption": "OAuth používá scope: Mail.ReadWrite, MailboxSettings.ReadWrite, Calendars.ReadWrite, offline_access.",
        "msal_not_installed": "MSAL není nainstalovaný. Doinstaluj závislosti z requirements.txt.",
        "graph_diag_header": "### Diagnostika Graph tokenu",
        "not_in_token": "(není v tokenu)",
        "graph_permissions_caption": (
            "Pro masterCategories je potřeba MailboxSettings.ReadWrite a pro kalendář Calendars.ReadWrite. "
            "Po změně oprávnění vždy vygeneruj nový token."
        ),
        "labels_header": "### Štítky pro AI kategorie",
        "use_custom_labels_check": "Použít vlastní mapování štítků",
        "use_custom_labels_help": "Když je vypnuto, použijí se původní MailAI štítky.",
        "load_outlook_labels_btn": "Načíst Outlook štítky",
        "outlook_labels_loaded": "Načteno Outlook štítků: {n}",
        "outlook_labels_error": "Nepodařilo se načíst Outlook štítky: {e}",
        "outlook_labels_caption": "Můžeš použít existující Outlook štítky nebo zadat vlastní názvy.",
        "custom_label_option": "(Vlastní štítek)",
        "custom_label_name": "Vlastní název",
        "add_deadline_label_check": "Přidávat doplňkový štítek pro termín",
        "deadline_label_selector": "Štítek pro termín",
        "deadline_label_custom_name": "Vlastní název termínového štítku",
        "default_labels_caption": "Výchozí režim: používají se původní MailAI štítky.",
        "custom_model_option": "(Vlastní model)",
        "model_label": "Model",
        "custom_model_name": "Vlastní název modelu",
        # Main page
        "inbox_section": "### Inbox souhrn (nepřečtené e-maily)",
        "analyze_btn": "Analyzovat nepřečtené e-maily za posledních N dní",
        "prompt_invalid_error": "Systémový prompt neprošel validací. Uprav ho před spuštěním analýzy.",
        "no_api_key_error": "Zadej LLM API key",
        "no_graph_token_error2": "Zadej Graph Access Token nebo se přihlas přes OAuth",
        "no_model_error": "Zadej nebo vyber model",
        "loading_emails_spinner": "Načítám e-maily z Microsoft Graph...",
        "not_replied_loaded": "Načteno e-mailů bez odpovědi: {n}",
        "unread_loaded": "Načteno nepřečtených e-mailů: {n}",
        "filter_stats_info": "Po filtraci štítků do analýzy: {n} | přeskočeno již označených: {m}",
        "urgent_reincluded_info": "Urgentní připomenutí vráceno do analýzy: {n}",
        "no_emails_warning": "Po filtraci nezbyly žádné e-maily pro analýzu.",
        "llm_limit_warning": "Pro LLM analýzu používám prvních {max_n} e-mailů z {total} kvůli rychlosti a stabilitě.",
        "analyzing_spinner": "Analyzuji přes LLM...",
        "analysis_done": "Souhrn hotový",
        "llm_timeout_error": (
            "LLM timeout: provider neodpověděl včas. Zkus jiný model, ověř Base URL, "
            "nebo zvýšit timeout v nastavení."
        ),
        "model_rejected_error": (
            "Model odmítl formát vstupu/výstupu. Zkus chat model (např. gpt-4o-mini) "
            "a klikni znovu na Načíst modely."
        ),
        "error_detail": "Detail chyby: {msg}",
        "generic_error": "Chyba: {e}",
        "result_header": "### Výsledek",
        "current_split_header": "### Aktuální rozdělení (možno upravit)",
        "current_split_caption": "Kategorie můžeš měnit přímo u každého e-mailu v tomto přehledu.",
        "deadlines_header": "### Termíny do kalendáře",
        "deadlines_caption": "U e-mailů s termínem můžeš jedním klikem vytvořit událost v Outlook kalendáři.",
        "no_hint": "bez upřesnění",
        "no_subject": "(bez předmětu)",
        "unknown_sender": "(neznámý odesílatel)",
        "add_to_calendar_btn": "Vložit do kalendáře",
        "calendar_event_body": "MailAI termín z e-mailu",
        "calendar_event_from": "Od:",
        "calendar_event_subject_lbl": "Předmět:",
        "calendar_event_reason": "Důvod:",
        "calendar_event_title_prefix": "Termín:",
        "calendar_event_created": "Událost vytvořena pro: {subject}",
        "calendar_event_error": "Nepodařilo se vytvořit událost v kalendáři: {e}",
        "bulk_actions_header": "### Doporučené hromadné akce",
        "mark_read_count": "Označit jako přečtené: {n}",
        "delete_count": "Smazat: 0 (zakázáno politikou aplikace)",
        "assign_labels_btn": "Přiřadit štítky podle aktuálního rozdělení",
        "categories_forbidden_warning": (
            "Graph token nemá oprávnění pro správu kategorií (masterCategories). "
            "Přidej scope MailboxSettings.ReadWrite a vygeneruj nový token. "
            "Pokusím se pokračovat: pokud kategorie už existují, přiřazení může fungovat."
        ),
        "categories_prepare_error": "Nepodařilo se připravit Outlook kategorie: {e}",
        "labels_assigned_success": "Štítek přiřazen u {ok} e-mailů",
        "labels_assign_fail": "Nepodařilo se přiřadit štítek u {fail} e-mailů",
        "categories_missing_info": (
            "Kategorie pravděpodobně v mailboxu neexistují. Po přidání oprávnění "
            "MailboxSettings.ReadWrite je aplikace vytvoří automaticky."
        ),
        "mark_read_btn": "Provést doporučené označení jako přečtené",
        "marked_read_success": "Označeno jako přečtené: {ok}",
        "mark_read_fail": "Nepovedlo se označit: {fail}",
        "permissions_caption": (
            "Pro hromadné akce je potřeba Mail.ReadWrite. Pro vytváření Outlook kategorií je potřeba "
            "MailboxSettings.ReadWrite. Pro vložení termínu do kalendáře je potřeba Calendars.ReadWrite."
        ),
        # Bucket labels
        "bucket_urgentni": "Urgentní",
        "bucket_stredne_dulezite": "Středně důležité",
        "bucket_pocka": "Počká",
        "bucket_k_preposlani": "K přeposlání",
        "bucket_ignorovat": "Ignorovat",
        # render_bucket
        "no_items": "Bez položek",
        "received_prefix": "doručeno:",
        "originally_prefix": "Původně:",
        "target_category": "Cílová kategorie",
        "outlook_link_title": "Primárně otevře v Outlook aplikaci, při nedostupnosti přejde na web",
        # validate_system_prompt
        "prompt_empty": "Systémový prompt je prázdný.",
        "prompt_missing_json": "Chybí explicitní požadavek na validní JSON výstup.",
        "prompt_missing_buckets": "Chybí definice pole 'buckets' pro rozdělení e-mailů.",
        "prompt_missing_bulk_actions": "Chybí definice pole 'recommended_bulk_actions'.",
        "prompt_missing_categories": "Chybí některé povinné kategorie: {cats}",
        # merge_prompt
        "merge_user_instructions": "Doplňující instrukce uživatele:",
        "merge_priority_senders": "Preferovaní odesílatelé:",
    },
    "en": {
        # Language selector
        "language_label": "Jazyk / Language",
        # Page
        "page_subtitle": "Local alternative without the Outlook add-in",
        # Sidebar
        "sidebar_header": "Settings",
        "system_prompt_expander": "Show system prompt",
        "system_prompt_label": "System prompt",
        "system_prompt_help": "Main AI analysis instructions. Saved to local settings.",
        "reset_prompt_btn": "Reset default system prompt",
        "reset_prompt_help": "Resets the system prompt to its original default.",
        "llm_timeout_label": "LLM timeout (seconds)",
        "analysis_mode_label": "Email selection mode",
        "analysis_mode_unread": "Unread",
        "analysis_mode_not_replied": "Not replied (Inbox vs Sent)",
        "label_mode_label": "Handling of already labeled emails",
        "label_mode_default": "Without MailAI label + urgent reminder",
        "label_mode_without": "Without MailAI label only",
        "label_mode_all": "All (including already labeled)",
        "urgent_reminder_label": "Remind urgent after (hours)",
        "calendar_tz_label": "Calendar timezone",
        "days_label": "Number of days back",
        "top_label": "Max number of emails",
        "custom_prompt_label": "Custom prompt",
        "custom_prompt_help": "Added to the default analysis instructions, does not replace them.",
        "priority_senders_label": "Priority senders",
        "auto_save_label": "Auto-save settings locally",
        "prompt_validation_warning": "System prompt validation found an issue that may break analysis:",
        "prompt_validation_ok": "System prompt passed basic validation.",
        "prompt_preview_expander": "Preview final prompt for the model",
        "save_btn": "Save",
        "clear_btn": "Clear",
        "settings_saved": "Settings saved to .mailai_local_settings.json",
        "settings_cleared": "Local saved settings cleared",
        "load_models_btn": "Load models",
        "models_loaded": "Loaded models: {n}",
        "models_load_error": "Error loading models: {e}",
        "verify_graph_btn": "Verify Graph permissions",
        "graph_token_label": "Graph Access Token (manual, optional)",
        "no_graph_token_error": "Please enter Graph Access Token or sign in via OAuth first",
        "oauth_header": "Microsoft OAuth (longer token lifetime)",
        "oauth_client_id_label": "Azure Client ID",
        "oauth_tenant_id_label": "Azure Tenant ID",
        "oauth_signin_btn": "Sign in with Microsoft",
        "oauth_signout_btn": "Sign out Microsoft account",
        "oauth_logged_in": "OAuth signed in: {account}",
        "oauth_logged_out": "OAuth not signed in",
        "oauth_missing_client_id": "Fill in Azure Client ID for OAuth.",
        "oauth_login_failed": "OAuth sign-in failed: {e}",
        "oauth_cache_cleared": "OAuth cache has been cleared.",
        "oauth_source_manual": "manual token",
        "oauth_source_oauth": "OAuth (MSAL cache)",
        "oauth_source_none": "not available",
        "oauth_token_source": "Graph token source: {source}",
        "oauth_scopes_caption": "OAuth uses scopes: Mail.ReadWrite, MailboxSettings.ReadWrite, Calendars.ReadWrite, offline_access.",
        "msal_not_installed": "MSAL is not installed. Install dependencies from requirements.txt.",
        "graph_diag_header": "### Graph Token Diagnostics",
        "not_in_token": "(not in token)",
        "graph_permissions_caption": (
            "MailboxSettings.ReadWrite is required for masterCategories and "
            "Calendars.ReadWrite is required for the calendar. "
            "Always generate a new token after changing permissions."
        ),
        "labels_header": "### Labels for AI categories",
        "use_custom_labels_check": "Use custom label mapping",
        "use_custom_labels_help": "When disabled, the default MailAI labels are used.",
        "load_outlook_labels_btn": "Load Outlook labels",
        "outlook_labels_loaded": "Loaded Outlook labels: {n}",
        "outlook_labels_error": "Failed to load Outlook labels: {e}",
        "outlook_labels_caption": "You can use existing Outlook labels or enter custom names.",
        "custom_label_option": "(Custom label)",
        "custom_label_name": "Custom name",
        "add_deadline_label_check": "Add additional label for deadline",
        "deadline_label_selector": "Label for deadline",
        "deadline_label_custom_name": "Custom deadline label name",
        "default_labels_caption": "Default mode: original MailAI labels are used.",
        "custom_model_option": "(Custom model)",
        "model_label": "Model",
        "custom_model_name": "Custom model name",
        # Main page
        "inbox_section": "### Inbox summary (unread emails)",
        "analyze_btn": "Analyze unread emails from the last N days",
        "prompt_invalid_error": "System prompt failed validation. Fix it before running analysis.",
        "no_api_key_error": "Enter LLM API key",
        "no_graph_token_error2": "Enter Graph Access Token or sign in via OAuth",
        "no_model_error": "Enter or select a model",
        "loading_emails_spinner": "Loading emails from Microsoft Graph...",
        "not_replied_loaded": "Loaded not-replied emails: {n}",
        "unread_loaded": "Loaded unread emails: {n}",
        "filter_stats_info": "After label filtering for analysis: {n} | skipped already labeled: {m}",
        "urgent_reincluded_info": "Urgent reminders returned to analysis: {n}",
        "no_emails_warning": "No emails remaining after filtering for analysis.",
        "llm_limit_warning": "Using the first {max_n} of {total} emails for LLM analysis for speed and stability.",
        "analyzing_spinner": "Analyzing via LLM...",
        "analysis_done": "Summary ready",
        "llm_timeout_error": (
            "LLM timeout: provider did not respond in time. Try a different model, "
            "check the Base URL, or increase the timeout in settings."
        ),
        "model_rejected_error": (
            "Model rejected the input/output format. Try a chat model (e.g. gpt-4o-mini) "
            "and click Load models again."
        ),
        "error_detail": "Error detail: {msg}",
        "generic_error": "Error: {e}",
        "result_header": "### Result",
        "current_split_header": "### Current categorization (editable)",
        "current_split_caption": "You can change categories directly for each email in this view.",
        "deadlines_header": "### Deadlines to calendar",
        "deadlines_caption": "For emails with a deadline, you can create an Outlook calendar event with one click.",
        "no_hint": "unspecified",
        "no_subject": "(no subject)",
        "unknown_sender": "(unknown sender)",
        "add_to_calendar_btn": "Add to calendar",
        "calendar_event_body": "MailAI deadline from email",
        "calendar_event_from": "From:",
        "calendar_event_subject_lbl": "Subject:",
        "calendar_event_reason": "Reason:",
        "calendar_event_title_prefix": "Deadline:",
        "calendar_event_created": "Event created for: {subject}",
        "calendar_event_error": "Failed to create calendar event: {e}",
        "bulk_actions_header": "### Recommended bulk actions",
        "mark_read_count": "Mark as read: {n}",
        "delete_count": "Delete: 0 (forbidden by application policy)",
        "assign_labels_btn": "Assign labels according to current categorization",
        "categories_forbidden_warning": (
            "Graph token does not have permission to manage categories (masterCategories). "
            "Add MailboxSettings.ReadWrite scope and generate a new token. "
            "I will try to continue: if categories already exist, assignment may work."
        ),
        "categories_prepare_error": "Failed to prepare Outlook categories: {e}",
        "labels_assigned_success": "Label assigned for {ok} emails",
        "labels_assign_fail": "Failed to assign label for {fail} emails",
        "categories_missing_info": (
            "Categories probably don't exist in the mailbox. After adding the "
            "MailboxSettings.ReadWrite permission, the application will create them automatically."
        ),
        "mark_read_btn": "Perform recommended mark as read",
        "marked_read_success": "Marked as read: {ok}",
        "mark_read_fail": "Failed to mark: {fail}",
        "permissions_caption": (
            "Mail.ReadWrite is required for bulk actions. "
            "MailboxSettings.ReadWrite is required for creating Outlook categories. "
            "Calendars.ReadWrite is required for adding deadlines to the calendar."
        ),
        # Bucket labels
        "bucket_urgentni": "Urgent",
        "bucket_stredne_dulezite": "Moderately important",
        "bucket_pocka": "Can wait",
        "bucket_k_preposlani": "To forward",
        "bucket_ignorovat": "Ignore",
        # render_bucket
        "no_items": "No items",
        "received_prefix": "received:",
        "originally_prefix": "Originally:",
        "target_category": "Target category",
        "outlook_link_title": "Opens in Outlook app primarily, falls back to web if unavailable",
        # validate_system_prompt
        "prompt_empty": "System prompt is empty.",
        "prompt_missing_json": "Missing explicit requirement for valid JSON output.",
        "prompt_missing_buckets": "Missing definition of 'buckets' field for email categorization.",
        "prompt_missing_bulk_actions": "Missing definition of 'recommended_bulk_actions' field.",
        "prompt_missing_categories": "Missing some required categories: {cats}",
        # merge_prompt
        "merge_user_instructions": "Additional user instructions:",
        "merge_priority_senders": "Priority senders:",
    },
}


def t(key: str, **kwargs) -> str:
    """Return the translated string for *key* in the currently active language."""
    lang = st.session_state.get("language", "cs")
    translations = TRANSLATIONS.get(lang, TRANSLATIONS["cs"])
    text = translations.get(key, TRANSLATIONS["cs"].get(key, key))
    if kwargs:
        try:
            text = text.format(**kwargs)
        except (KeyError, IndexError):
            pass
    return text


def default_bucket_label_map() -> dict[str, str]:
    return {bucket_key: MAILAI_CATEGORY_MAP[bucket_key][0] for bucket_key in BUCKET_ORDER}


def normalize_bucket_label_map(raw_map: dict | None) -> dict[str, str]:
    normalized = default_bucket_label_map()
    if not isinstance(raw_map, dict):
        return normalized

    for bucket_key in BUCKET_ORDER:
        value = str(raw_map.get(bucket_key, "") or "").strip()
        if value:
            normalized[bucket_key] = value
    return normalized


def load_local_settings() -> dict:
    if not SETTINGS_FILE.exists():
        return {}
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_local_settings(settings: dict) -> None:
    SETTINGS_FILE.write_text(json.dumps(settings, ensure_ascii=False, indent=2), encoding="utf-8")


def load_msal_cache():
    if msal is None:
        return None
    cache = msal.SerializableTokenCache()
    if MSAL_CACHE_FILE.exists():
        try:
            cache.deserialize(MSAL_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return cache


def save_msal_cache(cache) -> None:
    if cache is None:
        return
    if cache.has_state_changed:
        MSAL_CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")


def get_msal_app(client_id: str, tenant_id: str, cache):
    if msal is None or not client_id:
        return None
    authority = f"https://login.microsoftonline.com/{(tenant_id or 'common').strip() or 'common'}"
    return msal.PublicClientApplication(
        client_id=client_id.strip(),
        authority=authority,
        token_cache=cache,
    )


def acquire_graph_token_silent(client_id: str, tenant_id: str) -> tuple[str, str]:
    cache = load_msal_cache()
    app = get_msal_app(client_id, tenant_id, cache)
    if app is None:
        return "", ""

    accounts = app.get_accounts()
    if not accounts:
        return "", ""

    account = accounts[0]
    result = app.acquire_token_silent(MSAL_SCOPES, account=account)
    if isinstance(result, dict) and result.get("access_token"):
        save_msal_cache(cache)
        return str(result.get("access_token") or ""), str(account.get("username") or "")
    return "", ""


def acquire_graph_token_interactive(client_id: str, tenant_id: str) -> tuple[str, str, str]:
    cache = load_msal_cache()
    app = get_msal_app(client_id, tenant_id, cache)
    if app is None:
        return "", "", "MSAL not available or Client ID missing"

    try:
        result = app.acquire_token_interactive(scopes=MSAL_SCOPES)
    except Exception as e:
        return "", "", str(e)

    if isinstance(result, dict) and result.get("access_token"):
        save_msal_cache(cache)
        account = result.get("account") or {}
        account_name = str(account.get("username") or "")
        if not account_name:
            claims = result.get("id_token_claims") or {}
            account_name = str(claims.get("preferred_username") or claims.get("email") or "")
        return str(result.get("access_token") or ""), account_name, ""

    return "", "", str((result or {}).get("error_description") or (result or {}).get("error") or "Unknown error")


def clear_msal_auth_state() -> None:
    if MSAL_CACHE_FILE.exists():
        MSAL_CACHE_FILE.unlink()
    st.session_state.pop("graph_oauth_token", None)
    st.session_state.pop("graph_oauth_account", None)


def resolve_graph_token(manual_token: str, msal_client_id: str, msal_tenant_id: str) -> tuple[str, str]:
    manual_token = (manual_token or "").strip()
    if manual_token:
        return manual_token, t("oauth_source_manual")

    token, account_name = acquire_graph_token_silent(msal_client_id, msal_tenant_id)
    if token:
        st.session_state["graph_oauth_token"] = token
        if account_name:
            st.session_state["graph_oauth_account"] = account_name
        return token, t("oauth_source_oauth")

    return "", t("oauth_source_none")


def build_settings_payload() -> dict:
    return {
        "system_prompt": st.session_state.get("system_prompt", INBOX_PROMPT),
        "llm_api_key": st.session_state.get("llm_api_key", ""),
        "llm_base_url": st.session_state.get("llm_base_url", "https://llm.ai.e-infra.cz/v1/"),
        "llm_timeout": int(st.session_state.get("llm_timeout", 60)),
        "analysis_mode": st.session_state.get("analysis_mode", ANALYSIS_MODE_UNREAD),
        "label_handling_mode": st.session_state.get("label_handling_mode", LABEL_MODE_DEFAULT),
        "urgent_reminder_hours": int(st.session_state.get("urgent_reminder_hours", 24)),
        "model": st.session_state.get("model", ""),
        "graph_token": st.session_state.get("graph_token_input", ""),
        "msal_client_id": st.session_state.get("msal_client_id", DEFAULT_MS_CLIENT_ID),
        "msal_tenant_id": st.session_state.get("msal_tenant_id", DEFAULT_MS_TENANT_ID),
        "days": int(st.session_state.get("days", 10)),
        "top": int(st.session_state.get("top", 200)),
        "custom_prompt": st.session_state.get("custom_prompt", ""),
        "priority_senders_raw": st.session_state.get("priority_senders_raw", ""),
        "calendar_timezone": st.session_state.get("calendar_timezone", "Europe/Prague"),
        "use_custom_label_mapping": bool(st.session_state.get("use_custom_label_mapping", False)),
        "bucket_label_map": normalize_bucket_label_map(st.session_state.get("bucket_label_map")),
        "add_deadline_label": bool(st.session_state.get("add_deadline_label", True)),
        "deadline_label_name": st.session_state.get("deadline_label_name", MAILAI_DEADLINE_CATEGORY[0]),
        "auto_save_settings": bool(st.session_state.get("auto_save_settings", True)),
        "language": st.session_state.get("language", "cs"),
    }


def initialize_state_from_settings() -> None:
    if st.session_state.get("settings_initialized"):
        return
    saved = load_local_settings()
    st.session_state["system_prompt"] = saved.get("system_prompt", INBOX_PROMPT)
    st.session_state["llm_api_key"] = saved.get("llm_api_key", "")
    st.session_state["llm_base_url"] = saved.get("llm_base_url", "https://llm.ai.e-infra.cz/v1/")
    st.session_state["llm_timeout"] = int(saved.get("llm_timeout", 60))
    st.session_state["analysis_mode"] = saved.get("analysis_mode", ANALYSIS_MODE_UNREAD)
    st.session_state["label_handling_mode"] = saved.get("label_handling_mode", LABEL_MODE_DEFAULT)
    st.session_state["urgent_reminder_hours"] = int(saved.get("urgent_reminder_hours", 24))
    st.session_state["model"] = saved.get("model", "")
    st.session_state["graph_token_input"] = saved.get("graph_token", "")
    st.session_state["msal_client_id"] = saved.get("msal_client_id", DEFAULT_MS_CLIENT_ID)
    st.session_state["msal_tenant_id"] = saved.get("msal_tenant_id", DEFAULT_MS_TENANT_ID)
    st.session_state["days"] = int(saved.get("days", 10))
    st.session_state["top"] = int(saved.get("top", 200))
    st.session_state["custom_prompt"] = saved.get("custom_prompt", "")
    st.session_state["priority_senders_raw"] = saved.get("priority_senders_raw", "")
    st.session_state["calendar_timezone"] = saved.get("calendar_timezone", "Europe/Prague")
    st.session_state["use_custom_label_mapping"] = bool(saved.get("use_custom_label_mapping", False))
    st.session_state["bucket_label_map"] = normalize_bucket_label_map(saved.get("bucket_label_map"))
    st.session_state["add_deadline_label"] = bool(saved.get("add_deadline_label", True))
    st.session_state["deadline_label_name"] = str(saved.get("deadline_label_name", MAILAI_DEADLINE_CATEGORY[0]))
    st.session_state["auto_save_settings"] = bool(saved.get("auto_save_settings", True))
    st.session_state["language"] = saved.get("language", "cs")
    st.session_state["settings_initialized"] = True


def build_client(api_key: str, base_url: str, timeout_seconds: int = 60) -> OpenAI:
    # Keep retries low so blocked corporate networks fail fast with a clear message.
    return OpenAI(api_key=api_key, base_url=base_url, timeout=timeout_seconds, max_retries=1)


def merge_prompt(base_prompt: str, custom_prompt: str, senders: list[str]) -> str:
    parts = [base_prompt.strip()]

    custom_prompt = (custom_prompt or "").strip()
    if custom_prompt:
        parts.append(t("merge_user_instructions") + "\n" + custom_prompt)

    if senders:
        parts.append(t("merge_priority_senders") + " " + ", ".join(senders))

    return "\n\n".join(parts)


def reset_system_prompt() -> None:
    lang = st.session_state.get("language", "cs")
    st.session_state["system_prompt"] = INBOX_PROMPT_EN if lang == "en" else INBOX_PROMPT


def validate_system_prompt(prompt: str) -> list[str]:
    issues = []
    normalized_prompt = (prompt or "").strip()
    lowered = normalized_prompt.lower()

    if not normalized_prompt:
        return [t("prompt_empty")]

    if "json" not in lowered:
        issues.append(t("prompt_missing_json"))
    if "buckets" not in lowered:
        issues.append(t("prompt_missing_buckets"))
    if "recommended_bulk_actions" not in lowered:
        issues.append(t("prompt_missing_bulk_actions"))

    missing_categories = [bucket_key for bucket_key in BUCKET_ORDER if bucket_key not in lowered]
    if missing_categories:
        issues.append(t("prompt_missing_categories", cats=", ".join(missing_categories)))

    return issues


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
                "webLink": m.get("webLink", ""),
            }
        )
    return items


def _parse_graph_datetime(value: str) -> datetime:
    try:
        return datetime.fromisoformat((value or "").replace("Z", "+00:00"))
    except Exception:
        return datetime.now(timezone.utc)


def filter_items_for_analysis(items: list[dict], mode: str, urgent_reminder_hours: int) -> tuple[list[dict], dict]:
    if mode == LABEL_MODE_ALL:
        return items, {"skipped_labeled": 0, "urgent_reincluded": 0}

    unlabeled = []
    labeled = []
    for item in items:
        categories = item.get("categories") or []
        if any(str(cat).startswith(MAILAI_CATEGORY_PREFIX) for cat in categories):
            labeled.append(item)
        else:
            unlabeled.append(item)

    if mode == LABEL_MODE_WITHOUT:
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
            "$select": "id,conversationId,subject,from,receivedDateTime,bodyPreview,isRead,importance,categories,webLink",
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
            enriched_items.append(merged)
        buckets[bucket_key] = enriched_items

    result["buckets"] = buckets
    return result


def format_received_datetime(value: str) -> str:
    if not value:
        return ""
    try:
        dt = _parse_graph_datetime(value)
        if dt.tzinfo:
            dt = dt.astimezone()
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return ""


def build_outlook_app_link(web_link: str) -> str:
    if not web_link:
        return ""
    try:
        parsed = urlparse(web_link)
        if parsed.scheme not in ("http", "https"):
            return ""
        if "outlook" not in parsed.netloc and "office" not in parsed.netloc:
            return ""
        return urlunparse(("ms-outlook", parsed.netloc, parsed.path, parsed.params, parsed.query, parsed.fragment))
    except Exception:
        return ""


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

    deadline_label_name = (deadline_label_name or "").strip()
    if add_deadline_label and deadline_label_name and deadline_label_name not in existing:
        graph_create_master_category(token, deadline_label_name, MAILAI_DEADLINE_CATEGORY[1])


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
    label = t(f"bucket_{bucket_key}") if f"bucket_{bucket_key}" in TRANSLATIONS["cs"] else cfg["label"]
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
        st.caption(t("no_items"))
        return
    for itm in items[:50]:
        msg_id = str(itm.get("id") or "")
        subject = html.escape(str(itm.get("subject", "") or t("no_subject")))
        sender = html.escape(str(itm.get("from", "") or ""))
        web_link = str(itm.get("webLink") or "").strip()
        app_link = build_outlook_app_link(web_link)
        received_label = format_received_datetime(str(itm.get("receivedDateTime") or ""))
        received_suffix = f" | {t('received_prefix')} {received_label}" if received_label else ""
        link_title = t("outlook_link_title")
        if web_link and app_link:
            onclick_js = (
                f"event.preventDefault();window.location.href={json.dumps(app_link)};"
                f"setTimeout(function(){{window.location.href={json.dumps(web_link)};}},900);"
            )
            subject_html = (
                f'<a href="{html.escape(web_link, quote=True)}" target="_blank" '
                f'onclick="{html.escape(onclick_js, quote=True)}" '
                f'title="{html.escape(link_title, quote=True)}" '
                f'style="text-decoration:none;color:inherit"><strong>{subject}</strong></a>'
            )
        elif web_link:
            subject_html = (
                f'<a href="{html.escape(web_link, quote=True)}" target="_blank" '
                f'style="text-decoration:none;color:inherit"><strong>{subject}</strong></a>'
            )
        else:
            subject_html = f"<strong>{subject}</strong>"

        deadline_badge = ""
        if itm.get("has_deadline"):
            hint = itm.get("deadline_hint") or t("no_hint")
            deadline_badge = f' <span style="background:#7d3c98;color:#fff;border-radius:4px;padding:1px 6px;font-size:0.8rem">📅 {hint}</span>'
        moved_badge = ""
        if itm.get("suggested_bucket") and itm.get("suggested_bucket") != bucket_key:
            original_label = t(f"bucket_{itm['suggested_bucket']}") if f"bucket_{itm['suggested_bucket']}" in TRANSLATIONS["cs"] else BUCKET_UI.get(itm["suggested_bucket"], {}).get("label", itm["suggested_bucket"])
            moved_badge = (
                ' <span style="background:#ecf0f1;color:#2c3e50;border-radius:4px;padding:1px 6px;font-size:0.8rem">'
                f'{t("originally_prefix")} {original_label}</span>'
            )

        if editable and msg_id:
            col_info, col_choice = st.columns([5, 2])
            with col_info:
                st.markdown(
                    f'<span style="color:{color}">●</span> {subject_html} | '
                    f'<span style="color:#888">{sender}{received_suffix}</span>{deadline_badge}{moved_badge}',
                    unsafe_allow_html=True,
                )
                if itm.get("reason"):
                    st.caption(itm["reason"])
            with col_choice:
                st.selectbox(
                    t("target_category"),
                    options=list(BUCKET_ORDER),
                    format_func=lambda value: t(f"bucket_{value}") if f"bucket_{value}" in TRANSLATIONS["cs"] else BUCKET_UI[value]["label"],
                    key=f"bucket_override_{msg_id}",
                    label_visibility="collapsed",
                )
        else:
            st.markdown(
                f'<span style="color:{color}">●</span> {subject_html} | '
                f'<span style="color:#888">{sender}{received_suffix}</span>{deadline_badge}{moved_badge}',
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
    initialize_state_from_settings()
    st.caption(t("page_subtitle"))

    with st.sidebar:
        # Language selector – always at the top
        st.selectbox(
            t("language_label"),
            options=["cs", "en"],
            format_func=lambda v: "Čeština" if v == "cs" else "English",
            key="language",
        )

        st.header(t("sidebar_header"))

        # System prompt – collapsed by default to save space
        with st.expander(t("system_prompt_expander"), expanded=False):
            st.text_area(
                t("system_prompt_label"),
                height=220,
                key="system_prompt",
                help=t("system_prompt_help"),
            )
            st.button(
                t("reset_prompt_btn"),
                on_click=reset_system_prompt,
                help=t("reset_prompt_help"),
            )
        system_prompt = st.session_state.get("system_prompt", INBOX_PROMPT)

        llm_api_key = st.text_input("LLM API key", type="password", key="llm_api_key")
        llm_base_url = st.text_input("LLM Base URL", key="llm_base_url")
        llm_timeout = st.number_input(t("llm_timeout_label"), min_value=10, max_value=600, key="llm_timeout")

        st.subheader(t("oauth_header"))
        msal_client_id = st.text_input(t("oauth_client_id_label"), key="msal_client_id")
        msal_tenant_id = st.text_input(t("oauth_tenant_id_label"), key="msal_tenant_id")
        if msal is None:
            st.info(t("msal_not_installed"))
        else:
            if not st.session_state.get("oauth_bootstrap_done") and msal_client_id.strip():
                bootstrap_token, bootstrap_account = acquire_graph_token_silent(msal_client_id, msal_tenant_id)
                if bootstrap_token:
                    st.session_state["graph_oauth_token"] = bootstrap_token
                    st.session_state["graph_oauth_account"] = bootstrap_account
                st.session_state["oauth_bootstrap_done"] = True

            oauth_account = st.session_state.get("graph_oauth_account", "")
            if st.session_state.get("graph_oauth_token"):
                st.success(t("oauth_logged_in", account=oauth_account or "?"))
                if st.button(t("oauth_signout_btn")):
                    clear_msal_auth_state()
                    st.success(t("oauth_cache_cleared"))
                    st.rerun()
            else:
                st.caption(t("oauth_logged_out"))

            if st.button(t("oauth_signin_btn")):
                if not msal_client_id.strip():
                    st.error(t("oauth_missing_client_id"))
                else:
                    oauth_token, oauth_account, oauth_error = acquire_graph_token_interactive(msal_client_id, msal_tenant_id)
                    if oauth_token:
                        st.session_state["graph_oauth_token"] = oauth_token
                        st.session_state["graph_oauth_account"] = oauth_account
                        st.rerun()
                    else:
                        st.error(t("oauth_login_failed", e=oauth_error))
        st.caption(t("oauth_scopes_caption"))

        analysis_mode = st.selectbox(
            t("analysis_mode_label"),
            options=[ANALYSIS_MODE_UNREAD, ANALYSIS_MODE_NOT_REPLIED],
            format_func=lambda v: t("analysis_mode_unread") if v == ANALYSIS_MODE_UNREAD else t("analysis_mode_not_replied"),
            key="analysis_mode",
        )
        _label_mode_display = {
            LABEL_MODE_DEFAULT: "label_mode_default",
            LABEL_MODE_WITHOUT: "label_mode_without",
            LABEL_MODE_ALL: "label_mode_all",
        }
        label_handling_mode = st.selectbox(
            t("label_mode_label"),
            options=[LABEL_MODE_DEFAULT, LABEL_MODE_WITHOUT, LABEL_MODE_ALL],
            format_func=lambda v: t(_label_mode_display.get(v, v)),
            key="label_handling_mode",
        )
        urgent_reminder_hours = st.number_input(
            t("urgent_reminder_label"),
            min_value=1,
            max_value=240,
            key="urgent_reminder_hours",
        )
        graph_token = st.text_input(t("graph_token_label"), type="password", key="graph_token_input")
        effective_graph_token, token_source = resolve_graph_token(graph_token, msal_client_id, msal_tenant_id)
        st.caption(t("oauth_token_source", source=token_source))
        calendar_timezone = st.text_input(t("calendar_tz_label"), key="calendar_timezone")
        days = st.number_input(t("days_label"), min_value=1, max_value=30, key="days")
        top = st.number_input(t("top_label"), min_value=10, max_value=1000, key="top")
        custom_prompt = st.text_area(
            t("custom_prompt_label"),
            height=120,
            key="custom_prompt",
            help=t("custom_prompt_help"),
        )
        priority_senders_raw = st.text_area(t("priority_senders_label"), height=80, key="priority_senders_raw")
        auto_save_settings = st.checkbox(t("auto_save_label"), key="auto_save_settings")

        preview_senders = [
            sender.strip() for sender in priority_senders_raw.replace(",", "\n").split("\n") if sender.strip()
        ]
        prompt_validation_issues = validate_system_prompt(system_prompt)
        if prompt_validation_issues:
            st.warning(t("prompt_validation_warning"))
            for issue in prompt_validation_issues:
                st.caption(f"- {issue}")
        else:
            st.caption(t("prompt_validation_ok"))

        with st.expander(t("prompt_preview_expander")):
            st.code(merge_prompt(system_prompt or INBOX_PROMPT, custom_prompt, preview_senders), language="text")

        col_save, col_clear = st.columns(2)
        if col_save.button(t("save_btn")):
            save_local_settings(build_settings_payload())
            st.success(t("settings_saved"))
        if col_clear.button(t("clear_btn")):
            if SETTINGS_FILE.exists():
                SETTINGS_FILE.unlink()
            st.session_state["system_prompt"] = INBOX_PROMPT
            st.success(t("settings_cleared"))

        if st.button(t("load_models_btn")):
            try:
                client = build_client(llm_api_key, llm_base_url, int(llm_timeout))
                models = list_models(client)
                st.session_state["models"] = models
                st.success(t("models_loaded", n=len(models)))
            except Exception as e:
                st.error(t("models_load_error", e=e))

        if st.button(t("verify_graph_btn")):
            if not effective_graph_token:
                st.error(t("no_graph_token_error"))
            else:
                claims = decode_jwt_claims_unverified(effective_graph_token)
                scp = claims.get("scp", "")
                roles = claims.get("roles", [])
                aud = claims.get("aud", "")

                st.markdown(t("graph_diag_header"))
                st.write(f"aud: {aud}")
                st.write(f"scp: {scp or t('not_in_token')}")
                if roles:
                    st.write(f"roles: {roles}")

                me_status, me_msg = graph_endpoint_status(
                    effective_graph_token,
                    "https://graph.microsoft.com/v1.0/me",
                    {"$select": "id,userPrincipalName"},
                )
                msg_status, msg_msg = graph_endpoint_status(
                    effective_graph_token,
                    "https://graph.microsoft.com/v1.0/me/messages",
                    {"$top": 1, "$select": "id"},
                )
                cat_status, cat_msg = graph_endpoint_status(
                    effective_graph_token,
                    "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
                )
                evt_status, evt_msg = graph_endpoint_status(
                    effective_graph_token,
                    "https://graph.microsoft.com/v1.0/me/events",
                    {"$top": 1, "$select": "id"},
                )

                st.write(f"/me: {me_status} - {me_msg}")
                st.write(f"/me/messages: {msg_status} - {msg_msg}")
                st.write(f"/me/outlook/masterCategories: {cat_status} - {cat_msg}")
                st.write(f"/me/events: {evt_status} - {evt_msg}")
                st.caption(t("graph_permissions_caption"))

        st.markdown(t("labels_header"))
        use_custom_label_mapping = st.checkbox(
            t("use_custom_labels_check"),
            key="use_custom_label_mapping",
            help=t("use_custom_labels_help"),
        )
        if use_custom_label_mapping:
            if st.button(t("load_outlook_labels_btn")):
                if not effective_graph_token:
                    st.error(t("no_graph_token_error"))
                else:
                    try:
                        st.session_state["outlook_categories"] = sorted(graph_get_master_categories(effective_graph_token))
                        st.success(t("outlook_labels_loaded", n=len(st.session_state["outlook_categories"])))
                    except Exception as e:
                        st.error(t("outlook_labels_error", e=e))

            st.caption(t("outlook_labels_caption"))
            existing_labels = st.session_state.get("outlook_categories", [])
            saved_map = normalize_bucket_label_map(st.session_state.get("bucket_label_map"))
            updated_map = {}

            _custom_lbl = t("custom_label_option")
            for bucket_key in BUCKET_ORDER:
                display_name = t(f"bucket_{bucket_key}")
                current_label = saved_map[bucket_key]
                options = [_custom_lbl] + existing_labels if existing_labels else [_custom_lbl]
                default_idx = existing_labels.index(current_label) + 1 if current_label in existing_labels else 0
                selected_choice = st.selectbox(
                    display_name,
                    options=options,
                    index=default_idx,
                    key=f"bucket_label_choice_{bucket_key}",
                )

                if selected_choice == _custom_lbl:
                    custom_label = st.text_input(
                        t("custom_label_name"),
                        value=current_label,
                        key=f"bucket_label_custom_{bucket_key}",
                        label_visibility="collapsed",
                    )
                    updated_map[bucket_key] = custom_label.strip()
                else:
                    updated_map[bucket_key] = selected_choice

            st.session_state["bucket_label_map"] = normalize_bucket_label_map(updated_map)

            add_deadline_label = st.checkbox(
                t("add_deadline_label_check"),
                key="add_deadline_label",
            )
            if add_deadline_label:
                current_deadline_label = str(st.session_state.get("deadline_label_name", MAILAI_DEADLINE_CATEGORY[0]))
                deadline_options = [_custom_lbl] + existing_labels if existing_labels else [_custom_lbl]
                deadline_default_idx = existing_labels.index(current_deadline_label) + 1 if current_deadline_label in existing_labels else 0
                deadline_selected_choice = st.selectbox(
                    t("deadline_label_selector"),
                    options=deadline_options,
                    index=deadline_default_idx,
                    key="deadline_label_choice",
                )
                if deadline_selected_choice == _custom_lbl:
                    st.session_state["deadline_label_name"] = st.text_input(
                        t("deadline_label_custom_name"),
                        value=current_deadline_label,
                        key="deadline_label_custom",
                    ).strip()
                else:
                    st.session_state["deadline_label_name"] = deadline_selected_choice
        else:
            st.caption(t("default_labels_caption"))

        models = st.session_state.get("models", [])
        current_model = st.session_state.get("model", "")
        _custom_model = t("custom_model_option")
        if models:
            model_options = [_custom_model] + models
            default_model_idx = models.index(current_model) + 1 if current_model in models else 0
            selected_choice = st.selectbox(
                t("model_label"),
                options=model_options,
                index=default_model_idx,
            )
            if selected_choice == _custom_model:
                custom_model_value = current_model if current_model not in models else ""
                chosen_model = st.text_input(t("custom_model_name"), value=custom_model_value, key="model_custom")
            else:
                chosen_model = selected_choice
        else:
            chosen_model = st.text_input(t("model_label"), value=current_model, key="model_custom_no_list")

        if chosen_model != current_model:
            st.session_state["model"] = chosen_model

    final_model = st.session_state.get("model", "")

    st.markdown(t("inbox_section"))
    if st.button(t("analyze_btn"), type="primary"):
        prompt_issues = validate_system_prompt(system_prompt)
        if prompt_issues:
            st.error(t("prompt_invalid_error"))
            for issue in prompt_issues:
                st.caption(f"- {issue}")
            return
        if not llm_api_key:
            st.error(t("no_api_key_error"))
            return
        if not effective_graph_token:
            st.error(t("no_graph_token_error2"))
            return
        if not final_model:
            st.error(t("no_model_error"))
            return

        senders = [s.strip() for s in priority_senders_raw.replace(",", "\n").split("\n") if s.strip()]
        prompt = merge_prompt(system_prompt or INBOX_PROMPT, custom_prompt, senders)

        try:
            if auto_save_settings:
                save_local_settings(build_settings_payload())

            with st.spinner(t("loading_emails_spinner")):
                if analysis_mode == ANALYSIS_MODE_NOT_REPLIED:
                    items = fetch_not_replied_messages(effective_graph_token, int(days), int(top))
                    st.info(t("not_replied_loaded", n=len(items)))
                else:
                    items = fetch_unread_messages(effective_graph_token, int(days), int(top))
                    st.info(t("unread_loaded", n=len(items)))

            items, filter_stats = filter_items_for_analysis(
                items,
                label_handling_mode,
                int(urgent_reminder_hours),
            )
            if label_handling_mode != LABEL_MODE_ALL:
                st.info(t("filter_stats_info", n=len(items), m=filter_stats["skipped_labeled"]))
                if filter_stats["urgent_reincluded"]:
                    st.info(t("urgent_reincluded_info", n=filter_stats["urgent_reincluded"]))

            if not items:
                st.warning(t("no_emails_warning"))
                return

            llm_items = items[:MAX_EMAILS_FOR_LLM]
            if len(items) > MAX_EMAILS_FOR_LLM:
                st.warning(t("llm_limit_warning", max_n=MAX_EMAILS_FOR_LLM, total=len(items)))

            with st.spinner(t("analyzing_spinner")):
                client = build_client(llm_api_key, llm_base_url, int(llm_timeout))
                result = summarize_unread(client, final_model, prompt, llm_items, int(days))
                result = enforce_no_delete_policy(result)
                result = enrich_result_with_source_metadata(result, items)

            st.session_state["inbox_result"] = result
            st.session_state["graph_token"] = effective_graph_token
            initialize_bucket_overrides(result)
            st.success(t("analysis_done"))
        except APITimeoutError:
            st.error(t("llm_timeout_error"))
        except Exception as e:
            msg = str(e)
            if "required attributes" in msg.lower() or "požadovaných atribut" in msg.lower():
                st.error(t("model_rejected_error"))
                st.caption(t("error_detail", msg=msg))
            else:
                st.error(t("generic_error", e=e))

    result = st.session_state.get("inbox_result")
    if result:
        token = st.session_state.get("graph_token", "")
        bucket_overrides = get_bucket_overrides(result)
        effective_buckets = build_effective_buckets(result, bucket_overrides)
        effective_counts = {bucket_key: len(effective_buckets[bucket_key]) for bucket_key in BUCKET_ORDER}

        st.markdown(t("result_header"))
        st.write(result.get("overview", ""))

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric(t("bucket_urgentni"), effective_counts.get("urgentni", 0))
        c2.metric(t("bucket_stredne_dulezite"), effective_counts.get("stredne_dulezite", 0))
        c3.metric(t("bucket_pocka"), effective_counts.get("pocka", 0))
        c4.metric(t("bucket_k_preposlani"), effective_counts.get("k_preposlani", 0))
        c5.metric(t("bucket_ignorovat"), effective_counts.get("ignorovat", 0))

        st.markdown(t("current_split_header"))
        st.caption(t("current_split_caption"))
        for bkey in BUCKET_ORDER:
            render_bucket(bkey, effective_buckets.get(bkey, []), editable=True)

        deadline_items = get_deadline_items(effective_buckets)
        if deadline_items:
            st.markdown(t("deadlines_header"))
            st.caption(t("deadlines_caption"))
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
                    hint = itm.get("deadline_hint") or t("no_hint")
                    st.markdown(
                        f"**{itm.get('subject', t('no_subject'))}**  \n"
                        f"<span style='color:#666'>{itm.get('from', t('unknown_sender'))}</span>"
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
                    if st.button(t("add_to_calendar_btn"), key=f"event_btn_{msg_id}"):
                        start_dt = datetime.combine(st.session_state[date_key], st.session_state[time_key])
                        end_dt = start_dt + timedelta(minutes=int(st.session_state[dur_key]))
                        body_text = (
                            f"{t('calendar_event_body')}\n"
                            f"{t('calendar_event_from')} {itm.get('from', '')}\n"
                            f"{t('calendar_event_subject_lbl')} {itm.get('subject', '')}\n"
                            f"Deadline hint: {itm.get('deadline_hint') or ''}\n"
                            f"{t('calendar_event_reason')} {itm.get('reason') or ''}\n"
                            f"Message ID: {msg_id}"
                        )
                        try:
                            graph_create_calendar_event(
                                token,
                                f"{t('calendar_event_title_prefix')} {itm.get('subject', t('no_subject'))}",
                                start_dt,
                                end_dt,
                                calendar_timezone,
                                body_text,
                            )
                            st.success(t("calendar_event_created", subject=itm.get("subject", t("no_subject"))))
                        except Exception as e:
                            st.error(t("calendar_event_error", e=e))

        st.markdown(t("bulk_actions_header"))
        actions = result.get("recommended_bulk_actions", {})
        mark_ids = actions.get("mark_read_ids", [])

        st.write(t("mark_read_count", n=len(mark_ids)))
        st.write(t("delete_count"))
        if token:
            if st.button(t("assign_labels_btn")):
                ok = 0
                fail = 0
                categories_prepared = False
                use_custom_label_mapping = bool(st.session_state.get("use_custom_label_mapping", False))
                if use_custom_label_mapping:
                    selected_bucket_labels = normalize_bucket_label_map(st.session_state.get("bucket_label_map"))
                    add_deadline_label = bool(st.session_state.get("add_deadline_label", True))
                    deadline_cat_name = str(
                        st.session_state.get("deadline_label_name", MAILAI_DEADLINE_CATEGORY[0])
                    ).strip()
                else:
                    selected_bucket_labels = default_bucket_label_map()
                    add_deadline_label = True
                    deadline_cat_name = MAILAI_DEADLINE_CATEGORY[0]
                try:
                    if use_custom_label_mapping:
                        ensure_selected_master_categories(
                            token,
                            selected_bucket_labels,
                            add_deadline_label,
                            deadline_cat_name,
                        )
                    else:
                        ensure_mailai_master_categories(token)
                    categories_prepared = True
                except Exception as e:
                    if is_master_categories_forbidden(e):
                        st.warning(t("categories_forbidden_warning"))
                    else:
                        st.error(t("categories_prepare_error", e=e))
                        st.stop()

                for bucket_key in BUCKET_ORDER:
                    category_name = selected_bucket_labels.get(bucket_key, "").strip()
                    if not category_name:
                        continue
                    for itm in effective_buckets.get(bucket_key, []):
                        msg_id = itm.get("id")
                        if not msg_id:
                            continue
                        try:
                            graph_assign_category(token, msg_id, category_name)
                            ok += 1
                        except Exception:
                            fail += 1
                        if (
                            add_deadline_label
                            and deadline_cat_name
                            and bucket_key in ("urgentni", "stredne_dulezite")
                            and itm.get("has_deadline")
                        ):
                            try:
                                graph_assign_category(token, msg_id, deadline_cat_name)
                            except Exception:
                                pass

                st.success(t("labels_assigned_success", ok=ok))
                if fail:
                    st.warning(t("labels_assign_fail", fail=fail))
                if not categories_prepared and ok == 0:
                    st.info(t("categories_missing_info"))

            if st.button(t("mark_read_btn")):
                ok = 0
                fail = 0
                for msg_id in mark_ids:
                    try:
                        graph_patch_read(token, msg_id)
                        ok += 1
                    except Exception:
                        fail += 1
                st.success(t("marked_read_success", ok=ok))
                if fail:
                    st.warning(t("mark_read_fail", fail=fail))

            st.caption(t("permissions_caption"))


if __name__ == "__main__":
    main()

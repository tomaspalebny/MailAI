const INBOX_PROMPT = `Jsi emailovy asistent. Zpracuj seznam neprectenych emailu a vrat pouze validni JSON s touto strukturou:
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
- Kazdy email zarad prave do jedne kategorie.
- Nikdy nenavrhuj mazani emailu ani akci smazat.
- U urgentni a stredne_dulezite nastav has_deadline=true, pokud email obsahuje konkretni termin.
- Odpovidej cesky.
`;

const BUCKETS = ["urgentni", "stredne_dulezite", "pocka", "k_preposlani", "ignorovat"];
const TAGS = {
  urgentni: { key: "mailai_urgentni", name: "MailAI/Urgentni", color: "#e74c3c" },
  stredne_dulezite: { key: "mailai_stredne", name: "MailAI/Stredne dulezite", color: "#e67e22" },
  pocka: { key: "mailai_pocka", name: "MailAI/Pocka", color: "#d4ac0d" },
  k_preposlani: { key: "mailai_preposlani", name: "MailAI/K preposlani", color: "#2980b9" },
  ignorovat: { key: "mailai_ignorovat", name: "MailAI/Ignorovat", color: "#95a5a6" },
  deadline: { key: "mailai_deadline", name: "MailAI/S terminem", color: "#7d3c98" }
};

let lastResult = null;

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function parseJsonContent(content) {
  try {
    return JSON.parse(content);
  } catch (_err) {
    const cleaned = String(content || "")
      .replace(/```json/gi, "")
      .replace(/```/g, "")
      .trim();
    return JSON.parse(cleaned);
  }
}

function normalizeResult(raw) {
  const result = raw && typeof raw === "object" ? raw : {};
  const normalized = {
    overview: String(result.overview || ""),
    counts: {},
    buckets: {},
    recommended_bulk_actions: {
      mark_read_ids: []
    }
  };

  for (const bucket of BUCKETS) {
    normalized.counts[bucket] = Number(result.counts?.[bucket] || 0);
    normalized.buckets[bucket] = Array.isArray(result.buckets?.[bucket]) ? result.buckets[bucket] : [];
  }

  const markRead = result.recommended_bulk_actions?.mark_read_ids;
  normalized.recommended_bulk_actions.mark_read_ids = Array.isArray(markRead)
    ? markRead.map((v) => String(v))
    : [];

  return normalized;
}

async function getUnreadMessages(maxItems, daysBack) {
  const query = { unread: true };
  const fromDate = new Date(Date.now() - daysBack * 24 * 60 * 60 * 1000);

  // Some Thunderbird versions support query.fromDate, fallback is post-filtering.
  query.fromDate = fromDate;

  const result = await browser.messages.query(query);
  const all = [];
  let page = result;

  while (page) {
    const pageMessages = Array.isArray(page.messages) ? page.messages : [];
    for (const m of pageMessages) {
      const msgDate = m.date ? new Date(m.date) : null;
      if (msgDate && msgDate >= fromDate) {
        all.push(m);
      }
      if (all.length >= maxItems) {
        return all.slice(0, maxItems);
      }
    }

    if (!page.id || all.length >= maxItems) {
      break;
    }
    page = await browser.messages.continueList(page.id);
  }

  return all.slice(0, maxItems);
}

function toLlmItems(messages) {
  return messages.map((m) => ({
    id: String(m.id),
    subject: m.subject || "(bez predmetu)",
    from: m.author || "",
    receivedDateTime: m.date ? new Date(m.date).toISOString() : "",
    bodyPreview: m.snippet || "",
    tags: Array.isArray(m.tags) ? m.tags : []
  }));
}

async function analyzeWithLlm(config, messages) {
  const baseUrl = String(config.llmBaseUrl || "https://llm.ai.e-infra.cz/v1/").replace(/\/$/, "");
  const url = `${baseUrl}/chat/completions`;
  const apiKey = String(config.llmApiKey || "").trim();
  const model = String(config.model || "").trim();
  const customPrompt = String(config.customPrompt || "").trim();
  const senders = String(config.prioritySenders || "")
    .split(/\r?\n|,/)
    .map((s) => s.trim())
    .filter(Boolean);

  if (!apiKey) {
    throw new Error("Chybi LLM API key.");
  }
  if (!model) {
    throw new Error("Chybi model.");
  }

  const promptParts = [INBOX_PROMPT.trim()];
  if (customPrompt) {
    promptParts.push(`Doplnujici instrukce uzivatele:\n${customPrompt}`);
  }
  if (senders.length) {
    promptParts.push(`Preferovani odesilatele: ${senders.join(", ")}`);
  }

  const body = {
    model,
    response_format: { type: "json_object" },
    temperature: 0.2,
    max_tokens: 3500,
    messages: [
      { role: "system", content: promptParts.join("\n\n") },
      {
        role: "user",
        content: JSON.stringify(
          {
            window_days: Number(config.daysBack || 10),
            total_unread_fetched: messages.length,
            emails: toLlmItems(messages)
          },
          null,
          0
        )
      }
    ]
  };

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify(body)
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`LLM chyba ${response.status}: ${text.slice(0, 500)}`);
  }

  const json = await response.json();
  const content = json?.choices?.[0]?.message?.content;
  if (!content) {
    throw new Error("LLM nevratilo obsah odpovedi.");
  }

  return normalizeResult(parseJsonContent(content));
}

async function ensureTag(tagInfo) {
  const tags = await browser.messages.listTags();
  const existing = tags.find((t) => t.key === tagInfo.key);
  if (existing) {
    return existing.key;
  }

  await browser.messages.createTag(tagInfo.key, tagInfo.name, tagInfo.color);
  return tagInfo.key;
}

async function addTagToMessage(messageId, tagKey) {
  const msg = await browser.messages.get(messageId);
  const current = Array.isArray(msg.tags) ? msg.tags : [];
  if (current.includes(tagKey)) {
    return;
  }
  await browser.messages.update(messageId, { tags: [...current, tagKey] });
}

async function applyResult(config) {
  if (!lastResult) {
    throw new Error("Nejdriv spust analyzu.");
  }

  const addDeadlineTag = Boolean(config.addDeadlineTag);
  const tagKeys = {};

  for (const bucket of BUCKETS) {
    tagKeys[bucket] = await ensureTag(TAGS[bucket]);
  }
  if (addDeadlineTag) {
    tagKeys.deadline = await ensureTag(TAGS.deadline);
  }

  let tagged = 0;
  let markRead = 0;

  for (const bucket of BUCKETS) {
    const items = lastResult.buckets[bucket] || [];
    for (const item of items) {
      const id = Number(item.id);
      if (!Number.isFinite(id)) {
        continue;
      }

      await addTagToMessage(id, tagKeys[bucket]);
      tagged += 1;

      if (
        addDeadlineTag &&
        (bucket === "urgentni" || bucket === "stredne_dulezite") &&
        item.has_deadline
      ) {
        await addTagToMessage(id, tagKeys.deadline);
      }
    }
  }

  const markReadIds = lastResult.recommended_bulk_actions.mark_read_ids || [];
  for (const idValue of markReadIds) {
    const id = Number(idValue);
    if (!Number.isFinite(id)) {
      continue;
    }
    await browser.messages.update(id, { read: true });
    markRead += 1;
  }

  return { tagged, markRead };
}

browser.runtime.onMessage.addListener((message) => {
  if (!message || typeof message !== "object") {
    return null;
  }

  if (message.type === "mailai:saveConfig") {
    return browser.storage.local.set({ config: message.payload || {} });
  }

  if (message.type === "mailai:getConfig") {
    return browser.storage.local.get("config");
  }

  if (message.type === "mailai:analyze") {
    return (async () => {
      const config = message.payload || {};
      const maxItems = clamp(Number(config.maxItems || 120), 10, 500);
      const daysBack = clamp(Number(config.daysBack || 10), 1, 30);

      const unread = await getUnreadMessages(maxItems, daysBack);
      if (!unread.length) {
        lastResult = normalizeResult({ overview: "Zadne neprectene emaily.", counts: {}, buckets: {} });
        return { ok: true, unreadCount: 0, result: lastResult };
      }

      const analyzed = await analyzeWithLlm({ ...config, daysBack }, unread);
      const byId = new Map(unread.map((m) => [String(m.id), m]));
      for (const bucket of BUCKETS) {
        analyzed.buckets[bucket] = analyzed.buckets[bucket].map((item) => {
          const source = byId.get(String(item.id));
          return {
            ...item,
            subject: item.subject || source?.subject || "(bez predmetu)",
            from: item.from || source?.author || ""
          };
        });
        analyzed.counts[bucket] = analyzed.buckets[bucket].length;
      }

      lastResult = analyzed;
      return { ok: true, unreadCount: unread.length, result: analyzed };
    })();
  }

  if (message.type === "mailai:apply") {
    return (async () => {
      const applied = await applyResult(message.payload || {});
      return { ok: true, ...applied };
    })();
  }

  if (message.type === "mailai:lastResult") {
    return Promise.resolve({ result: lastResult });
  }

  return null;
});

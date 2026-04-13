const DEFAULT_CONFIG = {
  llmApiKey: "",
  llmBaseUrl: "https://llm.ai.e-infra.cz/v1/",
  model: "",
  daysBack: 10,
  maxItems: 120,
  customPrompt: "",
  prioritySenders: "",
  addDeadlineTag: true
};

const BUCKET_LABELS = {
  urgentni: "Urgentni",
  stredne_dulezite: "Stredne dulezite",
  pocka: "Pocka",
  k_preposlani: "K preposlani",
  ignorovat: "Ignorovat"
};

function $(id) {
  return document.getElementById(id);
}

function setStatus(text) {
  $("status").textContent = text;
}

function getConfigFromUi() {
  return {
    llmApiKey: $("llmApiKey").value.trim(),
    llmBaseUrl: $("llmBaseUrl").value.trim(),
    model: $("model").value.trim(),
    daysBack: Number($("daysBack").value || 10),
    maxItems: Number($("maxItems").value || 120),
    customPrompt: $("customPrompt").value,
    prioritySenders: $("prioritySenders").value,
    addDeadlineTag: $("addDeadlineTag").checked
  };
}

function setConfigToUi(config) {
  const merged = { ...DEFAULT_CONFIG, ...(config || {}) };
  $("llmApiKey").value = merged.llmApiKey;
  $("llmBaseUrl").value = merged.llmBaseUrl;
  $("model").value = merged.model;
  $("daysBack").value = String(merged.daysBack);
  $("maxItems").value = String(merged.maxItems);
  $("customPrompt").value = merged.customPrompt;
  $("prioritySenders").value = merged.prioritySenders;
  $("addDeadlineTag").checked = Boolean(merged.addDeadlineTag);
}

function renderResult(result) {
  const root = $("result");
  if (!result) {
    root.innerHTML = "";
    return;
  }

  const parts = [];
  parts.push(`<div><strong>Souhrn:</strong> ${escapeHtml(result.overview || "")}</div>`);

  for (const [key, label] of Object.entries(BUCKET_LABELS)) {
    const items = Array.isArray(result.buckets?.[key]) ? result.buckets[key] : [];
    parts.push(`<div class=\"bucket\"><div class=\"bucket-title\">${label}: ${items.length}</div>`);
    for (const item of items.slice(0, 20)) {
      parts.push(
        `<div class=\"item\"><div><strong>${escapeHtml(item.subject || "")}</strong></div><div>${escapeHtml(item.from || "")}</div><div>${escapeHtml(item.reason || "")}</div></div>`
      );
    }
    parts.push("</div>");
  }

  root.innerHTML = parts.join("");
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

async function loadConfig() {
  const data = await browser.runtime.sendMessage({ type: "mailai:getConfig" });
  setConfigToUi(data?.config || DEFAULT_CONFIG);
}

async function saveConfig() {
  const config = getConfigFromUi();
  await browser.runtime.sendMessage({ type: "mailai:saveConfig", payload: config });
  setStatus("Nastaveni ulozeno.");
}

async function analyze() {
  const config = getConfigFromUi();
  setStatus("Analyzuji neprectene emaily...");

  try {
    const response = await browser.runtime.sendMessage({ type: "mailai:analyze", payload: config });
    if (!response?.ok) {
      throw new Error(response?.error || "Neznama chyba");
    }

    renderResult(response.result);
    setStatus(`Hotovo. Nacteno ${response.unreadCount} emailu.`);
    await browser.runtime.sendMessage({ type: "mailai:saveConfig", payload: config });
  } catch (error) {
    setStatus(`Chyba: ${error.message || error}`);
  }
}

async function apply() {
  const config = getConfigFromUi();
  setStatus("Aplikuji tagy a mark as read...");

  try {
    const response = await browser.runtime.sendMessage({ type: "mailai:apply", payload: config });
    if (!response?.ok) {
      throw new Error(response?.error || "Neznama chyba");
    }
    setStatus(`Hotovo. Otagovano: ${response.tagged}, oznaceno jako prectene: ${response.markRead}.`);
  } catch (error) {
    setStatus(`Chyba: ${error.message || error}`);
  }
}

window.addEventListener("DOMContentLoaded", async () => {
  $("saveBtn").addEventListener("click", saveConfig);
  $("analyzeBtn").addEventListener("click", analyze);
  $("applyBtn").addEventListener("click", apply);

  await loadConfig();

  const cached = await browser.runtime.sendMessage({ type: "mailai:lastResult" });
  renderResult(cached?.result || null);
  setStatus("Pripraveno.");
});

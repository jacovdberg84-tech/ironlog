// IRONLOG/web/app/documents.js — Translations, AI documents, dark mode.
// Part of the main app; index.html loads these files in order and they share one global scope.

function applyI18n() {
  const map = {
    docsTitle: t("docsTitle"),
    docsSubtitle: t("docsSubtitle"),
    docsHeaderTitle: t("docsHeaderTitle"),
    docsDraftTitle: t("docsDraftTitle"),
  };
  Object.entries(map).forEach(([id, text]) => {
    const el = qs(id);
    if (el) el.textContent = text;
  });
  const statusEl = qs("status");
  if (statusEl && (!statusEl.textContent || statusEl.textContent.trim() === "Ready.")) {
    statusEl.textContent = t("statusReady");
  }
}

async function saveDocHeader() {
  const payload = {
    name: qs("docHeaderName")?.value || "",
    site_name: qs("docHeaderSite")?.value || "",
    department: qs("docHeaderDepartment")?.value || "",
    prepared_by: qs("docHeaderPreparedBy")?.value || "",
    approved_by: qs("docHeaderApprovedBy")?.value || "",
    revision: qs("docHeaderRevision")?.value || "",
  };
  const res = await fetchJson(`${API}/api/docs/headers`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setStatus(`Header saved (#${res.id}).`);
  await loadDocHeaders();
}

async function loadDocHeaders() {
  const listEl = qs("docHeaderList");
  if (!listEl) return;
  const res = await fetchJson(`${API}/api/docs/headers`);
  if (!res.rows?.length) {
    listEl.innerHTML = `<div class="item"><small>No header profiles yet.</small></div>`;
    return;
  }
  listEl.innerHTML = "";
  res.rows.forEach((r) => {
    const d = document.createElement("div");
    d.className = "item";
    d.innerHTML = `<b>#${r.id} ${r.name}</b> - ${r.site_name || "-"} / ${r.department || "-"} <small>(Rev: ${r.revision || "-"})</small>`;
    d.addEventListener("click", () => {
      const idEl = qs("docHeaderId");
      if (idEl) idEl.value = String(r.id);
    });
    listEl.appendChild(d);
  });
  const idEl = qs("docHeaderId");
  if (idEl && !Number(idEl.value || 0) && res.rows[0]?.id) {
    idEl.value = String(res.rows[0].id);
  }
}

async function generateDocDraft() {
  const payload = {
    header_id: Number(qs("docHeaderId")?.value || 0),
    doc_type: qs("docType")?.value || "SOP",
    title: qs("docTitle")?.value || "",
    language: qs("docLanguage")?.value || getLang(),
    scope: qs("docScope")?.value || "",
    hazards: qs("docHazards")?.value || "",
    controls: qs("docControls")?.value || "",
    extra_notes: qs("docInputs")?.value || "",
  };
  const res = await fetchJson(`${API}/api/docs/draft-generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const out = qs("docDraftOutput");
  if (out) out.textContent = res.draft_text || "";
  const idEl = qs("docDraftId");
  if (idEl) idEl.value = String(res.id || "");
  setStatus(`Draft generated (#${res.id}).`);
  await loadDocDrafts();
}

async function generateDocDraftFromRequest(requestArg) {
  const requestText = String(requestArg || qs("aiSmartPrompt")?.value || "").trim();
  if (!requestText) return alert("Enter what document you want first.");
  let headerId = Number(qs("docHeaderId")?.value || 0);
  if (!headerId) {
    try {
      const hdr = await fetchJson(`${API}/api/docs/headers`);
      if (Array.isArray(hdr.rows) && hdr.rows[0]?.id) {
        headerId = Number(hdr.rows[0].id);
        const idEl = qs("docHeaderId");
        if (idEl) idEl.value = String(headerId);
      }
    } catch (_) {
      // Backend will still fallback to latest header when possible.
    }
  }
  const payload = {
    request_text: requestText,
    header_id: headerId,
    doc_type: qs("docType")?.value || "SOP",
    title: qs("docTitle")?.value || requestText.slice(0, 80),
    language: qs("docLanguage")?.value || getLang(),
    scope: qs("docScope")?.value || "",
    hazards: qs("docHazards")?.value || "",
    controls: qs("docControls")?.value || "",
  };
  const res = await fetchJson(`${API}/api/docs/draft-generate-request`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const out = qs("docDraftOutput");
  if (out) {
    const ctx = Array.isArray(res.related_docs) && res.related_docs.length
      ? `\n\n[Related docs used]\n${res.related_docs.map((d) => `#${d.id} ${d.title}`).join("\n")}`
      : "";
    out.textContent = (res.draft_text || "") + ctx;
  }
  const idEl = qs("docDraftId");
  if (idEl) idEl.value = String(res.id || "");
  setStatus(`Draft generated from request (#${res.id})${res.ai_used ? " with AI" : " (template fallback)"}.`);
  await loadDocDrafts();
}

function inferDocTypeFromPrompt(prompt) {
  const s = String(prompt || "").toLowerCase();
  if (s.includes("checklist")) return "Checklist";
  if (s.includes("method statement")) return "Method Statement";
  if (s.includes("site instruction")) return "Site Instruction";
  if (s.includes("risk")) return "Risk Note";
  if (s.includes("sop") || s.includes("procedure")) return "SOP";
  return "";
}

function parseMachineProblemFromPrompt(prompt) {
  const src = String(prompt || "").trim();
  if (!src) return { machine: "", problem: "" };
  const m = src.match(/^(.+?)\s+(?:has|have|with|showing|shows|no)\s+(.+)$/i);
  if (m) {
    const machine = String(m[1] || "").replace(/\s+$/, "").trim();
    const problem = src.slice(machine.length).replace(/^\s*(has|have|with|showing|shows)?\s*/i, "").trim();
    return { machine, problem };
  }
  return { machine: "", problem: src };
}

async function runAiSmart() {
  const out = qs("askJakesOutput");
  const smartPrompt = String(qs("aiSmartPrompt")?.value || "").trim();

  if (!smartPrompt) {
    alert("Enter a question or document request first.");
    return;
  }

  if (out) {
    out.textContent = "⏳ Thinking...";
    out.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }

  const faultKeywords = /(fault|error|not working|no\s+hydraulic|no\s+hydraulics|no\s+start|won't start|wont start|leak|overheat|pressure|engine|starter|battery|transmission|brake)/i;
  const docKeywords = /(sop|checklist|method statement|site instruction|risk note|document|procedure|policy|template)/i;

  try {
    if (faultKeywords.test(smartPrompt) && !docKeywords.test(smartPrompt)) {
      const parsed = parseMachineProblemFromPrompt(smartPrompt);
      await askJakes({ machine: parsed.machine, problem: parsed.problem, context: "" });
      return;
    }

    const inferred = inferDocTypeFromPrompt(smartPrompt);
    const titleEl = qs("docTitle");
    if (titleEl && !String(titleEl.value || "").trim()) titleEl.value = smartPrompt.slice(0, 80);
    if (inferred) {
      const typeEl = qs("docType");
      if (typeEl) typeEl.value = inferred;
    }
    await generateDocDraftFromRequest(smartPrompt);
    if (out) {
      out.textContent = "Document draft generated — see the Draft Output box below.";
    }
  } catch (err) {
    if (out) out.textContent = "❌ Error: " + (err.message || String(err));
    setStatus("Smart AI error: " + err.message);
  }
}

function applyAskJakesPreset(type) {
  const machineEl = qs("askJakesMachine");
  const problemEl = qs("askJakesProblem");
  const contextEl = qs("askJakesContext");
  if (!machineEl || !problemEl || !contextEl) return;

  if (type === "hydraulics") {
    machineEl.value = machineEl.value || "CAT 950 Loader";
    problemEl.value = "No hydraulics";
    contextEl.value = "Engine starts, steering weak, no bucket lift.";
    return;
  }
  if (type === "starting") {
    machineEl.value = machineEl.value || "CAT 950 Loader";
    problemEl.value = "Will not start";
    contextEl.value = "Battery indicator low, starter clicking.";
    return;
  }
  if (type === "overheat") {
    machineEl.value = machineEl.value || "CAT 950 Loader";
    problemEl.value = "Engine overheating";
    contextEl.value = "Temperature rises under load, fan noise normal.";
  }
}

function useAskJakesAnswerAsNotes() {
  const answer = String(qs("askJakesOutput")?.textContent || "").trim();
  if (!answer) {
    alert("Ask Jakes first to get an answer.");
    return;
  }
  const notesEl = qs("docInputs");
  if (!notesEl) return;
  const existing = String(notesEl.value || "").trim();
  notesEl.value = existing ? `${existing}\n\nAsk Jakes notes:\n${answer}` : `Ask Jakes notes:\n${answer}`;
  setStatus("Ask Jakes answer copied to draft notes.");
}

async function askJakes(override = {}) {
  const machine = String(override.machine || "").trim();
  const problem = String(override.problem || "").trim();
  const context = String(override.context || "").trim();
  const out = qs("askJakesOutput");

  if (!problem) {
    if (out) out.textContent = "❌ Please describe the machine problem.";
    return;
  }

  try {
    const res = await fetchJson(`${API}/api/docs/ai/ask`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ machine, problem, context }),
    });

    if (out) {
      out.textContent = String(res.answer || "No answer returned.");
      out.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }
    setStatus("Ask Jakes answered.");
  } catch (err) {
    if (out) out.textContent = "❌ Error: " + (err.message || String(err));
    setStatus("Ask Jakes error: " + err.message);
  }
}

function speakDocDraft() {
  const text = String(qs("docDraftOutput")?.textContent || "").trim();
  if (!text) {
    alert("Generate a draft first.");
    return;
  }
  if (!("speechSynthesis" in window)) {
    alert("Speech is not supported in this browser.");
    return;
  }
  const lang = String(qs("docLanguage")?.value || getLang() || "en").toLowerCase();
  const voiceLang = lang === "af" ? "af-ZA" : lang === "zu" ? "zu-ZA" : lang === "pt" ? "pt-PT" : "en-US";
  window.speechSynthesis.cancel();
  const utter = new SpeechSynthesisUtterance(text.slice(0, 12000));
  utter.lang = voiceLang;
  utter.rate = 1;
  utter.pitch = 1;
  window.speechSynthesis.speak(utter);
  setStatus("Speaking draft...");
}

function stopSpeakingDocDraft() {
  if ("speechSynthesis" in window) {
    window.speechSynthesis.cancel();
    setStatus("Speech stopped.");
  }
}

async function openDocDraftPdf(download = false) {
  const id = Number(qs("docDraftId")?.value || 0);
  if (!id) return alert("Enter/select a Draft ID first.");
  try {
    const check = await fetchJson(`${API}/api/docs/drafts`);
    const row = Array.isArray(check.rows) ? check.rows.find((r) => Number(r.id) === id) : null;
    const decision = String(row?.decision || "").toLowerCase();
    if (decision !== "approved") {
      setStatus(`Draft #${id} is '${decision || "pending"}'. Approve it (Yes) before PDF export.`);
      alert("Only approved documents can be exported to PDF.");
      return;
    }
    const url = `${API}/api/docs/drafts/${id}.pdf${download ? "?download=1" : ""}`;
    await openAuthedReport(url, { download, filename: `IRONLOG_Document_Draft_${id}.pdf` });
  } catch (e) {
    setStatus("Open PDF failed: " + (e.message || e));
  }
}

function openDocRegisterPdf(download = false) {
  const currentOnly = Boolean(qs("docRegisterCurrentOnly")?.checked);
  const params = new URLSearchParams();
  if (download) params.set("download", "1");
  if (currentOnly) params.set("current_only", "1");
  const q = params.toString();
  const url = `${API}/api/docs/register.pdf${q ? `?${q}` : ""}`;
  return openAuthedReport(url, { download, filename: "IRONLOG_Document_Register.pdf" });
}

async function openDocDraftWord(download = false) {
  const id = Number(qs("docDraftId")?.value || 0);
  if (!id) return alert("Enter/select a Draft ID first.");
  try {
    const check = await fetchJson(`${API}/api/docs/drafts`);
    const row = Array.isArray(check.rows) ? check.rows.find((r) => Number(r.id) === id) : null;
    const decision = String(row?.decision || "").toLowerCase();
    if (decision !== "approved") {
      setStatus(`Draft #${id} is '${decision || "pending"}'. Approve it (Yes) before Word export.`);
      alert("Only approved documents can be exported to Word.");
      return;
    }
    const url = `${API}/api/docs/drafts/${id}.docx${download ? "?download=1" : ""}`;
    await downloadAuthedFile(url, `IRONLOG_Document_Draft_${id}.docx`);
  } catch (e) {
    setStatus("Open Word failed: " + (e.message || e));
  }
}

function openDocRegisterWord(download = false) {
  const currentOnly = Boolean(qs("docRegisterCurrentOnly")?.checked);
  const params = new URLSearchParams();
  if (download) params.set("download", "1");
  if (currentOnly) params.set("current_only", "1");
  const q = params.toString();
  const url = `${API}/api/docs/register.docx${q ? `?${q}` : ""}`;
  return downloadAuthedFile(url, "IRONLOG_Document_Register.docx");
}

async function decideDocDraft(approved) {
  const id = Number(qs("docDraftId")?.value || 0);
  if (!id) return alert("Enter/select a Draft ID first.");
  const res = await fetchJson(`${API}/api/docs/drafts/${id}/decision`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ approved }),
  });
  setStatus(`Draft #${res.id} marked ${res.decision}.`);
  await loadDocDrafts();
}

async function loadDocDrafts() {
  const listEl = qs("docDraftsList");
  if (!listEl) return;
  const res = await fetchJson(`${API}/api/docs/drafts`);
  if (!res.rows?.length) {
    listEl.innerHTML = `<div class="item"><small>No drafts yet.</small></div>`;
    return;
  }
  listEl.innerHTML = "";
  const rows = Array.isArray(res.rows) ? res.rows : [];
  const supersededByApproved = new Set(
    rows
      .filter((x) => String(x?.decision || "").toLowerCase() === "approved" && Number(x?.supersedes_draft_id || 0) > 0)
      .map((x) => Number(x.supersedes_draft_id))
  );
  const currentOnly = Boolean(qs("docDraftsCurrentOnly")?.checked);
  const visibleRows = currentOnly
    ? rows.filter((r) => {
      const decision = String(r?.decision || "").toLowerCase();
      return decision === "approved" && !supersededByApproved.has(Number(r?.id || 0));
    })
    : rows;
  if (!visibleRows.length) {
    listEl.innerHTML = `<div class="item"><small>${currentOnly ? "No current approved drafts." : "No drafts yet."}</small></div>`;
    return;
  }
  visibleRows.forEach((r) => {
    const d = document.createElement("div");
    d.className = "item";
    const decision = String(r.decision || "").toLowerCase();
    const canPdf = decision === "approved";
    const isCurrentApproved = decision === "approved" && !supersededByApproved.has(Number(r.id));
    const stateBadge = isCurrentApproved
      ? `<span class="pill blue">Current</span>`
      : `<span class="pill orange">Historical</span>`;
    const rev = r.revision_no ? `Rev ${r.revision_no}` : "Rev -";
    const supersedes = r.supersedes_draft_id ? ` | supersedes #${r.supersedes_draft_id}` : "";
    d.innerHTML = `<b>#${r.id}</b> ${r.doc_type} - ${r.title} <small>[${r.language}]</small> <span class="pill">${r.decision}</span> ${stateBadge} <small>${rev}${supersedes} | Header: ${r.header_name || "-"}</small> <button data-doc-open-pdf="${r.id}" ${canPdf ? "" : "disabled"} title="${canPdf ? "Open final PDF" : "Approve first"}">Open PDF</button>`;
    d.addEventListener("click", () => {
      const idEl = qs("docDraftId");
      if (idEl) idEl.value = String(r.id);
    });
    d.querySelector("button[data-doc-open-pdf]")?.addEventListener("click", (evt) => {
      evt.stopPropagation();
      const idEl = qs("docDraftId");
      if (idEl) idEl.value = String(r.id);
      openDocDraftPdf(false);
    });
    listEl.appendChild(d);
  });
}

function initDarkMode() {
  const toggle = document.getElementById("darkModeToggle");
  if (!toggle) return;
  
  const saved = localStorage.getItem("ironlog-theme");
  if (saved === "dark") {
    document.documentElement.setAttribute("data-theme", "dark");
  }
  
  toggle.addEventListener("click", () => {
    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    if (isDark) {
      document.documentElement.removeAttribute("data-theme");
      localStorage.setItem("ironlog-theme", "light");
    } else {
      document.documentElement.setAttribute("data-theme", "dark");
      localStorage.setItem("ironlog-theme", "dark");
    }
  });
}

/** Start-up: AI document header, draft and Ask Jakes controls. Called once from init() in init.js. */
function wireDocumentControls() {
  qs("saveDocHeaderBtn")?.addEventListener("click", () =>
    saveDocHeader().catch((e) => setStatus("Header save error: " + e.message))
  );
  qs("loadDocHeadersBtn")?.addEventListener("click", () =>
    loadDocHeaders().catch((e) => setStatus("Header load error: " + e.message))
  );
  qs("generateDocDraftBtn")?.addEventListener("click", () =>
    generateDocDraft().catch((e) => setStatus("Draft generate error: " + e.message))
  );
  qs("generateDocDraftFromRequestBtn")?.addEventListener("click", () =>
    generateDocDraftFromRequest().catch((e) => setStatus("Draft request generate error: " + e.message))
  );
  qs("aiSmartRunBtn")?.addEventListener("click", () =>
    runAiSmart().catch((e) => setStatus("Smart AI error: " + e.message))
  );
  qs("askJakesBtn")?.addEventListener("click", () =>
    askJakes().catch((e) => setStatus("Ask Jakes error: " + e.message))
  );
  qs("askJakesPresetHydraulics")?.addEventListener("click", () =>
    applyAskJakesPreset("hydraulics")
  );
  qs("askJakesPresetStarting")?.addEventListener("click", () =>
    applyAskJakesPreset("starting")
  );
  qs("askJakesPresetOverheat")?.addEventListener("click", () =>
    applyAskJakesPreset("overheat")
  );
  qs("askJakesUseAsNotesBtn")?.addEventListener("click", () =>
    useAskJakesAnswerAsNotes()
  );
  qs("speakDocDraftBtn")?.addEventListener("click", () =>
    speakDocDraft()
  );
  qs("stopSpeakDocDraftBtn")?.addEventListener("click", () =>
    stopSpeakingDocDraft()
  );
  qs("docApproveYesBtn")?.addEventListener("click", () =>
    decideDocDraft(true).catch((e) => setStatus("Draft decision error: " + e.message))
  );
  qs("docApproveNoBtn")?.addEventListener("click", () =>
    decideDocDraft(false).catch((e) => setStatus("Draft decision error: " + e.message))
  );
  qs("openDocDraftPdfBtn")?.addEventListener("click", () =>
    openDocDraftPdf(false)
  );
  qs("downloadDocDraftPdfBtn")?.addEventListener("click", () =>
    openDocDraftPdf(true)
  );
  qs("openDocDraftWordBtn")?.addEventListener("click", () =>
    openDocDraftWord(false)
  );
  qs("downloadDocDraftWordBtn")?.addEventListener("click", () =>
    openDocDraftWord(true)
  );
  qs("openDocRegisterPdfBtn")?.addEventListener("click", () =>
    openDocRegisterPdf(false)
  );
  qs("downloadDocRegisterPdfBtn")?.addEventListener("click", () =>
    openDocRegisterPdf(true)
  );
  qs("openDocRegisterWordBtn")?.addEventListener("click", () =>
    openDocRegisterWord(false)
  );
  qs("downloadDocRegisterWordBtn")?.addEventListener("click", () =>
    openDocRegisterWord(true)
  );
  qs("loadDocDraftsBtn")?.addEventListener("click", () =>
    loadDocDrafts().catch((e) => setStatus("Draft list error: " + e.message))
  );
  qs("docDraftsCurrentOnly")?.addEventListener("change", () =>
    loadDocDrafts().catch((e) => setStatus("Draft list error: " + e.message))
  );
}

/** Start-up: Initial AI document lists. Called once from init() in init.js. */
function loadDocumentsOnStartup() {
  loadDocHeaders().catch(() => {});
  loadDocDrafts().catch(() => {});
}

// IRONLOG/web/app/borris.js — Borris (Ironmind) insights, reports and Q&A.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function loadIronmindInsight(options = {}) {
  const silent = Boolean(options.silent);
  const summaryEl = qs("ironmindSummary");
  const metaEl = qs("ironmindMeta");

  try {
    const res = await fetchJson(`${API}/api/ironmind/latest?report_type=daily_admin`);
    const report = res?.report || null;
    if (!report) {
      if (summaryEl) {
        const emptySections = parseIronmindSections(
          ["BORRIS DAILY INSIGHT", "", "Repairs Needed", "- Insufficient data",
           "", "Operational Risks", "- Insufficient data",
           "", "Suggestions", "- Insufficient data",
           "", "Data Gaps", "- Insufficient data",
           "", "Data Anomalies", "- Insufficient data"].join("\n")
        );
        renderIronmindSections(summaryEl, emptySections);
      }
      if (metaEl) metaEl.textContent = "No report generated yet.";
      if (!silent) setStatus("BORRIS insight not available yet.");
      return;
    }

    renderIronmindReport(report);
    if (!silent) setStatus("BORRIS insight loaded.");
  } catch (err) {
    if (metaEl) metaEl.textContent = "Insight unavailable right now.";
    if (!silent) setStatus("BORRIS load error: " + err.message);
    throw err;
  }
}

async function loadIronmindHealth() {
  const healthEl = qs("ironmindHealth");
  if (!healthEl) return;
  try {
    const data = await fetchJson(`${API}/api/ironmind/health`);
    const live = Boolean(data?.live_enabled);
    const provider = String(data?.provider || "none");
    const model = String(data?.model || "");
    const mode = String(data?.last_ask_mode || "unknown");
    const err = String(data?.last_ask_error || "").trim();
    const left = live ? `OpenAI: Connected (${provider}${model ? `/${model}` : ""})` : `OpenAI: Not connected (${provider})`;
    const right = `Last ask mode: ${mode}${err ? ` | Last error: ${err}` : ""}`;
    healthEl.textContent = `${left} | ${right}`;
    healthEl.className = live ? "status-ok" : "status-overdue";
  } catch (e) {
    healthEl.textContent = `OpenAI status unavailable: ${e.message || e}`;
    healthEl.className = "status-overdue";
  }
}

async function loadIronmindSettings() {
  const meta = qs("ironmindSettingsMeta");
  try {
    const res = await fetchJson(`${API}/api/ironmind/settings`);
    const s = res?.settings || {};
    const setVal = (id, v) => {
      const el = qs(id);
      if (el && Number.isFinite(Number(v))) el.value = Number(v);
    };
    setVal("ironmindMaxRunGlobal", s.max_daily_run_hours);
    setVal("ironmindMaxRunLdv", s.max_daily_run_hours_ldv);
    setVal("ironmindMaxRunTruck", s.max_daily_run_hours_truck);
    setVal("ironmindMaxRunHeavy", s.max_daily_run_hours_heavy);
    if (meta) meta.textContent = "Thresholds loaded.";
  } catch {
    if (meta) meta.textContent = "Threshold load failed.";
  }
}

async function saveIronmindSettings() {
  const meta = qs("ironmindSettingsMeta");
  const readNum = (id, d) => {
    const n = Number(qs(id)?.value);
    return Number.isFinite(n) && n > 0 ? n : d;
  };
  const payload = {
    max_daily_run_hours: readNum("ironmindMaxRunGlobal", 24),
    max_daily_run_hours_ldv: readNum("ironmindMaxRunLdv", 16),
    max_daily_run_hours_truck: readNum("ironmindMaxRunTruck", 18),
    max_daily_run_hours_heavy: readNum("ironmindMaxRunHeavy", 24),
  };
  setStatus("Saving Borris thresholds...");
  try {
    await fetchJson(`${API}/api/ironmind/settings`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (meta) meta.textContent = "Saved. Click Refresh Insight to apply now.";
    setStatus("Borris thresholds saved.");
  } catch (err) {
    if (meta) meta.textContent = `Save failed: ${err.message || err}`;
    setStatus("Borris threshold save failed.");
  }
}

function summarizeIronmindText(text) {
  const oneLine = String(text || "")
    .replace(/\s+/g, " ")
    .replace(/^(?:BORRIS|IRONMIND) DAILY INSIGHT\s*/i, "")
    .trim();
  if (!oneLine) return "No summary text.";
  return oneLine.length > 150 ? `${oneLine.slice(0, 147)}...` : oneLine;
}

function toYmd(date) {
  const d = new Date(date);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function buildRecentYmds(days) {
  const out = [];
  const now = new Date();
  for (let i = 0; i < days; i += 1) {
    const d = new Date(now);
    d.setDate(now.getDate() - i);
    out.push(toYmd(d));
  }
  return out;
}

function renderIronmindReport(report) {
  const summaryEl = qs("ironmindSummary");
  const metaEl = qs("ironmindMeta");
  if (summaryEl) {
    const sections = parseIronmindSections(String(report?.summary || ""));
    renderIronmindSections(summaryEl, sections);
  }
  if (metaEl) {
    const created = report?.created_at ? String(report.created_at).replace("T", " ").slice(0, 16) : "-";
    metaEl.textContent = `Report date: ${report?.report_date || "-"} | Updated: ${created}`;
  }
}

async function loadIronmindHistory(options = {}) {
  const silent = Boolean(options.silent);
  const listEl = qs("ironmindHistoryList");
  const includeMissing = Boolean(qs("ironmindShowMissingDays")?.checked);
  if (!listEl) return;
  try {
    const res = await fetchJson(`${API}/api/ironmind/history?report_type=daily_admin&days=7`);
    const rowsRaw = Array.isArray(res?.reports) ? res.reports : [];
    const rowsByDate = new Map(rowsRaw.map((r) => [String(r.report_date || "").trim(), r]));
    const rows = includeMissing
      ? buildRecentYmds(7).map((ymd) => {
          if (rowsByDate.has(ymd)) return rowsByDate.get(ymd);
          return {
            id: 0,
            report_date: ymd,
            report_type: "daily_admin",
            created_at: "-",
            summary: "BORRIS DAILY INSIGHT\n\nRepairs Needed\n- Insufficient data\n\nOperational Risks\n- Insufficient data\n\nSuggestions\n- Insufficient data\n\nData Gaps\n- Insufficient data",
            synthetic_missing: true,
          };
        })
      : rowsRaw;
    listEl.innerHTML = "";
    if (!rows.length) {
      listEl.appendChild(item("<small>No BORRIS history yet.</small>"));
      if (!silent) setStatus("No BORRIS history found.");
      return;
    }

    rows.forEach((r) => {
      const created = r?.created_at && r.created_at !== "-" ? String(r.created_at).replace("T", " ").slice(0, 16) : "-";
      const preview = summarizeIronmindText(r?.summary || "");
      const previewClass = r?.synthetic_missing ? "ironmind-history-preview missing" : "ironmind-history-preview";
      const updatedText = r?.synthetic_missing ? "No generated report" : `Updated ${escapeHtml(created)}`;
      const node = item(
        `<div class="ironmind-history-item">` +
          `<div class="ironmind-history-meta"><b>${escapeHtml(r.report_date || "-")}</b> · ${updatedText}</div>` +
          `<div class="${previewClass}">${escapeHtml(preview)}</div>` +
          `<button class="ironmind-history-open" data-ironmind-history-id="${Number(r.id || 0)}">${r?.synthetic_missing ? "Open placeholder" : "Open report"}</button>` +
        `</div>`
      );
      node.dataset.ironmindRow = JSON.stringify(r);
      listEl.appendChild(node);
    });
    if (!silent) setStatus("BORRIS history loaded.");
  } catch (err) {
    listEl.innerHTML = `<small class="muted">History unavailable right now.</small>`;
    if (!silent) setStatus("BORRIS history error: " + err.message);
    throw err;
  }
}

async function refreshIronmindInsight() {
  const btn = qs("ironmindRefreshBtn");
  const contextNotes = String(qs("ironmindContext")?.value || "").trim();
  const detailMode = Boolean(qs("ironmindDetailMode")?.checked);
  if (btn) btn.disabled = true;
  setStatus("Refreshing BORRIS insight...");
  try {
    await fetchJson(`${API}/api/ironmind/run`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        force: true,
        report_date: todayLocalYmd(),
        report_type: "daily_admin",
        context_notes: contextNotes || undefined,
        detail_mode: detailMode,
      }),
    });
    await loadIronmindInsight({ silent: true });
    await loadIronmindHistory({ silent: true });
    setStatus("BORRIS insight refreshed.");
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function planIronmindWeek() {
  const button = qs('ironmindWeeklyPlanBtn');
  const out = qs('ironmindAskResult');
  if (button) button.disabled = true;
  if (out) out.textContent = 'Borris is preparing next week’s maintenance draft...';
  try {
    const plan = await fetchJson(`${API}/api/maintenance/weekly-plan`);
    if (!plan.ok) throw new Error(plan.error || 'Planning failed');
    if (out) { out.textContent = plan.short_answer; out.style.whiteSpace = 'pre-wrap'; }
    setStatus('Weekly maintenance draft ready for review.');
  } catch (err) {
    if (out) out.textContent = 'Weekly plan failed: ' + (err.message || err);
  } finally { if (button) button.disabled = false; }
}

async function askIronmindQuestion() {
  const input = qs("ironmindAskInput");
  const out = qs("ironmindAskResult");
  const date = qs("date")?.value || todayLocalYmd();
  const contextNotes = String(qs("ironmindContext")?.value || "").trim();
  window.__ironmindAskHistory = Array.isArray(window.__ironmindAskHistory) ? window.__ironmindAskHistory : [];
  const question = String(input?.value || "").trim();
  if (!question) {
    if (out) out.innerHTML = `<small class="muted">Type a question first.</small>`;
    return;
  }
  if (out) out.innerHTML = `<small class="muted">Asking BORRIS...</small>`;
  try {
    const res = await fetchJson(`${API}/api/ironmind/ask`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        question,
        date,
        history: window.__ironmindAskHistory.slice(-6),
        ...(contextNotes ? { context_notes: contextNotes } : {}),
      }),
    });
    const short = String(res?.short_answer || res?.answer || res?.message || "No answer returned.");
    const safe = escapeHtml(short).replace(/\n/g, "<br>");
    window.__ironmindAskHistory.push({ question, answer: short });
    window.__ironmindAskHistory = window.__ironmindAskHistory.slice(-12);
    saveIronmindAskHistoryLocal();
    renderIronmindAskHistory();
    const sid = getIronmindSessionId();
    fetchJson(`${API}/api/ironmind/ask-memory`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ session_id: sid, question, answer: short }),
    }).catch(() => {});
    loadIronmindHealth().catch(() => {});
    setStatus("BORRIS question answered.");
  } catch (e) {
    if (out) out.innerHTML = `<small class="muted">Question failed: ${escapeHtml(e.message || String(e))}</small>`;
    setStatus("BORRIS ask error: " + (e.message || e));
  }
}

function ironmindMemoryStorageKey() {
  return `ironmind_chat_${String(getSessionUser() || "user").toLowerCase()}_${String(getSessionSite() || "main").toLowerCase()}`;
}

function getIronmindSessionId() {
  const key = `${ironmindMemoryStorageKey()}_session`;
  let sid = String(localStorage.getItem(key) || "").trim();
  if (!sid) {
    sid = `sess-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
    localStorage.setItem(key, sid);
  }
  return sid;
}

function saveIronmindAskHistoryLocal() {
  try {
    localStorage.setItem(ironmindMemoryStorageKey(), JSON.stringify(window.__ironmindAskHistory || []));
  } catch {}
}

function loadIronmindAskHistoryLocal() {
  try {
    const parsed = JSON.parse(String(localStorage.getItem(ironmindMemoryStorageKey()) || "[]"));
    if (Array.isArray(parsed)) return parsed;
  } catch {}
  return [];
}

function renderIronmindAskHistory() {
  const out = qs("ironmindAskResult");
  if (!out) return;
  const rows = Array.isArray(window.__ironmindAskHistory) ? window.__ironmindAskHistory : [];
  if (!rows.length) {
    out.innerHTML = `<small class="muted">What are we working on today?</small>`;
    return;
  }
  out.innerHTML = rows.map((h) => `
    <div class="borris-message borris-message-user"><span class="borris-speaker">You</span><div>${escapeHtml(String(h.question || "")).replace(/\n/g, "<br>")}</div></div>
    <div class="borris-message borris-message-assistant"><span class="borris-speaker">Borris</span><div>${escapeHtml(String(h.answer || "")).replace(/\n/g, "<br>")}</div></div>
  `).join("");
  out.scrollTop = out.scrollHeight;
}

async function hydrateIronmindAskMemory() {
  window.__ironmindAskHistory = loadIronmindAskHistoryLocal().slice(-12);
  renderIronmindAskHistory();
  try {
    const sid = getIronmindSessionId();
    const data = await fetchJson(`${API}/api/ironmind/ask-memory?session_id=${encodeURIComponent(sid)}&limit=30`);
    const rows = Array.isArray(data?.rows) ? data.rows : [];
    if (rows.length) {
      window.__ironmindAskHistory = rows.map((r) => ({
        question: String(r.question || ""),
        answer: String(r.answer || ""),
      })).slice(-12);
      saveIronmindAskHistoryLocal();
      renderIronmindAskHistory();
    }
  } catch {}
}

async function resetIronmindAskMemory() {
  const out = qs("ironmindAskResult");
  const key = ironmindMemoryStorageKey();
  try {
    localStorage.removeItem(key);
  } catch {}
  window.__ironmindAskHistory = [];
  renderIronmindAskHistory();
  try {
    const sid = getIronmindSessionId();
    await fetchJson(`${API}/api/ironmind/ask-memory?session_id=${encodeURIComponent(sid)}`, {
      method: "DELETE",
    });
  } catch {}
  if (out) out.innerHTML = `<small class="muted">Memory cleared.</small>`;
  setStatus("Borris chat memory reset.");
}

async function generateIronmindRsgPlan(createWo = false) {
  const out = qs("ironmindRsgResult");
  const assetCode = String(qs("ironmindRsgAssetCode")?.value || "").trim().toUpperCase();
  const serviceHours = Number(qs("ironmindRsgHours")?.value || 2000);
  if (!assetCode) {
    if (out) out.textContent = "Asset code is required.";
    return;
  }
  const body = {
    asset_code: assetCode,
    service_hours: Number.isFinite(serviceHours) && serviceHours > 0 ? serviceHours : 2000,
  };
  if (out) out.textContent = "";
  setStatus(createWo ? "Creating RSG work order..." : "Generating RSG plan...");
  try {
    const endpoint = createWo ? "/api/ironmind/rsg/create-wo" : "/api/ironmind/rsg/plan";
    const res = await fetchJson(`${API}${endpoint}`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify(body),
    });
    if (out) out.textContent = JSON.stringify(res, null, 2);
    if (createWo && Number(res?.work_order_id || 0) > 0) {
      setStatus(`RSG WO created (#${Number(res.work_order_id)}).`);
      await loadDashboard().catch(() => {});
    } else {
      setStatus("RSG plan generated.");
    }
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus(`RSG ${createWo ? "WO creation" : "plan"} failed.`);
  }
}

async function previewIronmindRsgPdf() {
  const out = qs("ironmindRsgResult");
  const assetCode = String(qs("ironmindRsgAssetCode")?.value || "").trim().toUpperCase();
  const serviceHours = Number(qs("ironmindRsgHours")?.value || 2000);
  if (!assetCode) {
    if (out) out.textContent = "Asset code is required.";
    return;
  }
  const hours = Number.isFinite(serviceHours) && serviceHours > 0 ? serviceHours : 2000;
  setStatus("Opening RSG PDF preview...");
  try {
    const url = `${API}/api/ironmind/rsg/preview.pdf?asset_code=${encodeURIComponent(assetCode)}&service_hours=${encodeURIComponent(hours)}`;
    const opened = await openAuthedReport(url, { filename: `IRONLOG_RSG_${assetCode}_${hours}h.pdf` });
    if (opened !== false) setStatus("RSG PDF preview opened.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("RSG PDF preview failed.");
  }
}

async function downloadIronmindRsgPdf() {
  const out = qs("ironmindRsgResult");
  const assetCode = String(qs("ironmindRsgAssetCode")?.value || "").trim().toUpperCase();
  const serviceHours = Number(qs("ironmindRsgHours")?.value || 2000);
  if (!assetCode) {
    if (out) out.textContent = "Asset code is required.";
    return;
  }
  const hours = Number.isFinite(serviceHours) && serviceHours > 0 ? serviceHours : 2000;
  setStatus("Preparing RSG PDF download...");
  try {
    const url = `${API}/api/ironmind/rsg/preview.pdf?asset_code=${encodeURIComponent(assetCode)}&service_hours=${encodeURIComponent(hours)}&download=1`;
    const downloaded = await downloadAuthedFile(url, `IRONLOG_RSG_${assetCode}_${hours}h.pdf`);
    if (downloaded) setStatus("RSG PDF download started.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("RSG PDF download failed.");
  }
}

function parseIronmindSections(text) {
  const defs = [
    { key: "repairs", name: "Repairs Needed" },
    { key: "risks", name: "Operational Risks" },
    { key: "suggestions", name: "Suggestions" },
    { key: "data_gaps", name: "Data Gaps" },
    { key: "anomalies", name: "Data Anomalies" },
  ];
  const src = String(text || "");
  return defs.map((sec, i) => {
    const next = defs[i + 1];
    const start = src.indexOf(sec.name);
    if (start === -1) return { key: sec.key, name: sec.name, items: [] };
    const end = next ? src.indexOf(next.name, start + sec.name.length) : src.length;
    const block = src.slice(start + sec.name.length, end === -1 ? src.length : end);
    const items = block
      .split("\n")
      .map((l) => l.trim())
      .filter((l) => l.startsWith("-"))
      .map((l) => l.slice(1).trim())
      .filter(Boolean);
    return { key: sec.key, name: sec.name, items: items.length ? items : ["Insufficient data"] };
  });
}

function renderIronmindSections(summaryEl, sections) {
  if (!summaryEl) return;
  const navMap = {
    repairs: { label: "\u2192 View Assets", key: "repairs" },
    risks: { label: "\u2192 View Stock", key: "risks" },
  };
  // Pattern: UPPERCASE asset code at start of item before ":"
  const assetPat = /^([A-Z][A-Z0-9_-]{1,9}):\s/;
  let html = `<div class="ironmind-header-line">BORRIS DAILY INSIGHT</div>`;
  for (const sec of sections) {
    const nav = navMap[sec.key];
    const drillBtn = nav
      ? `<button class="ironmind-drill" data-ironmind-drill="${sec.key}">${escapeHtml(nav.label)}</button>`
      : "";
    html += `<div class="ironmind-title-row"><span class="ironmind-section-name">${escapeHtml(sec.name)}</span>${drillBtn}</div>`;
    html += `<ul class="ironmind-items">`;
    for (const itm of sec.items) {
      const m = sec.key === "repairs" ? itm.match(assetPat) : null;
      if (m) {
        const code = m[1];
        const rest = escapeHtml(itm.slice(m[0].length));
        html += `<li class="ironmind-item">- <button class="ironmind-asset-link" data-ironmind-asset="${escapeHtml(code)}">${escapeHtml(code)}</button>: ${rest}</li>`;
      } else {
        html += `<li class="ironmind-item">- ${escapeHtml(itm)}</li>`;
      }
    }
    html += `</ul>`;
  }
  summaryEl.innerHTML = html;
}

function ironmindDrillDown(sectionKey) {
  if (sectionKey === "repairs") {
    switchTab("assets");
  } else if (sectionKey === "risks") {
    switchTab("stock");
  }
}

async function ironmindGoToAsset(assetCode) {
  switchTab("assets");
  await loadAssetsFleet().catch(() => {});
  const code = String(assetCode || "").trim();
  if (!code) return;
  const exists = assetsFleetCache.some((c) => c.asset_code === code);
  if (exists) {
    await selectAssetCard(code, { loadHistory: true, scroll: true });
  }
}

/** Start-up: Borris insight, health, report and question controls (plus dashboard threshold save). Called once from init() in init.js. */
function wireBorrisControls() {
  qs("ironmindRefreshBtn")?.addEventListener("click", () =>
    refreshIronmindInsight().catch((e) => setStatus("BORRIS refresh error: " + e.message))
  );
  setInterval(() => {
    loadIronmindHealth().catch(() => {});
  }, 30000);
  qs("ironmindSaveSettingsBtn")?.addEventListener("click", () =>
    saveIronmindSettings().catch((e) => setStatus("Borris settings error: " + (e.message || e)))
  );
  qs("ironmindWeeklyPlanBtn")?.addEventListener("click", planIronmindWeek);
  qs("ironmindAskBtn")?.addEventListener("click", () =>
    askIronmindQuestion().catch((e) => setStatus("BORRIS ask error: " + e.message))
  );
  qs("ironmindResetMemoryBtn")?.addEventListener("click", () =>
    resetIronmindAskMemory().catch((e) => setStatus("BORRIS reset memory error: " + e.message))
  );
  qs("ironmindAskInput")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      askIronmindQuestion().catch((err) => setStatus("BORRIS ask error: " + err.message));
    }
  });
  hydrateIronmindAskMemory().catch(() => {});
  qs("saveThresholds")?.addEventListener("click", () => saveThresholdsFromUI());
  qs("saveLdvThresholds")?.addEventListener("click", () => saveLdvPrestartThresholdsFromUI());
  qs("ironmindSummary")?.addEventListener("click", (e) => {
    const el = e.target instanceof HTMLElement ? e.target : null;
    if (!el) return;
    const drillKey = el.dataset.ironmindDrill;
    const assetCode = el.dataset.ironmindAsset;
    if (drillKey) ironmindDrillDown(drillKey);
    if (assetCode) ironmindGoToAsset(assetCode).catch(() => {});
  });
  qs("ironmindHistoryList")?.addEventListener("click", (e) => {
    const el = e.target instanceof HTMLElement ? e.target.closest("button[data-ironmind-history-id]") : null;
    if (!el) return;
    const rowEl = el.closest(".item");
    if (!rowEl?.dataset?.ironmindRow) return;
    try {
      const row = JSON.parse(rowEl.dataset.ironmindRow);
      renderIronmindReport(row);
      setStatus(`Opened BORRIS report for ${row?.report_date || "-"}.`);
    } catch (_) {
      setStatus("Unable to open selected BORRIS report.");
    }
  });
  qs("riskBoardList")?.addEventListener("click", (e) => {
    const target = e.target instanceof HTMLElement ? e.target : null;
    if (!target) return;
    const openBtn = target.closest("button[data-ironmind-risk-asset]");
    if (openBtn) {
      const code = String(openBtn.getAttribute("data-ironmind-risk-asset") || "").trim();
      if (!code) return;
      ironmindGoToAsset(code).catch(() => {});
      return;
    }
    const woBtn = target.closest("button[data-ironmind-risk-wo]");
    if (woBtn) {
      const code = String(woBtn.getAttribute("data-ironmind-risk-wo") || "").trim();
      if (!code) return;
      (async () => {
        const downDesc = "BORRIS predicted risk work order";
        const res = await fetchJson(`${API}/api/breakdowns/ensure-open`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            asset_code: code,
            breakdown_date: date,
            description: downDesc,
            critical: false,
          }),
        });
        const woId = Number(res?.primary_work_order_id || 0);
        setStatus(woId > 0 ? `WO #${woId} ready for ${code}.` : `Work order ensured for ${code}.`);
        await loadDashboard().catch(() => {});
      })().catch((err) => setStatus("Create WO failed: " + (err.message || err)));
    }
  });
  qs("ironmindShowMissingDays")?.addEventListener("change", () => {
    loadIronmindHistory({ silent: true }).catch(() => {});
  });
  qs("ironmindReloadHistory")?.addEventListener("click", () => {
    loadIronmindHistory().catch((e) => setStatus("BORRIS history error: " + e.message));
  });
  qs("ironmindRsgPlanBtn")?.addEventListener("click", () => {
    generateIronmindRsgPlan(false).catch((e) => setStatus("RSG plan error: " + (e.message || e)));
  });
  qs("ironmindRsgPreviewPdfBtn")?.addEventListener("click", () => {
    previewIronmindRsgPdf().catch((e) => setStatus("RSG preview error: " + (e.message || e)));
  });
  qs("ironmindRsgDownloadPdfBtn")?.addEventListener("click", () => {
    downloadIronmindRsgPdf().catch((e) => setStatus("RSG download error: " + (e.message || e)));
  });
  qs("ironmindRsgCreateWoBtn")?.addEventListener("click", () => {
    generateIronmindRsgPlan(true).catch((e) => setStatus("RSG create WO error: " + (e.message || e)));
  });
}

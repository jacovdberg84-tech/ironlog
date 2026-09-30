// IRONLOG technician portal: Today, job, shift, week and machine screens.
// Talks to /api/tech (a thin layer over work orders, stock, parts requests and
// breakdowns). Writes made without signal are kept on the phone and sent later
// with the same client_event_id, so a retry never saves twice.
(function () {
  const A = window.IronlogAuth;
  if (!A) return;

  const QUEUE_KEY = "ironlog-tech-queue-v1";
  const CACHE_KEY = "ironlog-tech-cache-v1";
  const LEAD_ROLES = ["supervisor", "workshop_admin", "admin", "plant_manager", "site_manager"];
  const esc = (s) => A.escapeHtml(s == null ? "" : String(s));
  const qs = (id) => document.getElementById(id);
  const api = (p) => `${A.API}/api/tech${p}`;

  let root = null;
  let view = { name: "today" };
  let tick = null;
  let syncing = false;

  // ------------------------------------------------------------ small helpers
  function uid() {
    if (window.crypto?.randomUUID) return window.crypto.randomUUID();
    return `t${Date.now().toString(36)}${Math.random().toString(36).slice(2, 10)}`;
  }

  function read(key, fallback) {
    try {
      const v = JSON.parse(localStorage.getItem(key) || "null");
      return v == null ? fallback : v;
    } catch {
      return fallback;
    }
  }

  function write(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch {
      return false;
    }
  }

  function cacheGet(k) {
    return read(CACHE_KEY, {})[k] || null;
  }

  function cachePut(k, v) {
    const c = read(CACHE_KEY, {});
    c[k] = { at: new Date().toISOString(), data: v };
    const keys = Object.keys(c);
    if (keys.length > 40) keys.sort((a, b) => String(c[a].at).localeCompare(String(c[b].at))).slice(0, keys.length - 40).forEach((x) => delete c[x]);
    write(CACHE_KEY, c);
  }

  function isLead() {
    return A.getSessionRoles().some((r) => LEAD_ROLES.includes(String(r).toLowerCase()));
  }

  function time(iso) {
    if (!iso) return "";
    const d = new Date(String(iso).includes("T") ? iso : `${String(iso).replace(" ", "T")}Z`);
    if (Number.isNaN(d.getTime())) return String(iso);
    return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", hour12: false });
  }

  function dayTime(iso) {
    if (!iso) return "";
    const d = new Date(String(iso).includes("T") ? iso : `${String(iso).replace(" ", "T")}Z`);
    if (Number.isNaN(d.getTime())) return String(iso);
    return `${d.toLocaleDateString([], { day: "2-digit", month: "short" })} ${time(iso)}`;
  }

  function hrs(h) {
    const n = Number(h || 0);
    if (n < 1) return `${Math.round(n * 60)} min`;
    const hh = Math.floor(n);
    const mm = Math.round((n - hh) * 60);
    return mm ? `${hh} h ${mm} min` : `${hh} h`;
  }

  function since(iso) {
    const ms = Date.now() - Date.parse(iso);
    return Number.isFinite(ms) && ms > 0 ? hrs(ms / 3600000) : "0 min";
  }

  function toast(msg, kind = "") {
    const t = qs("tpToast");
    if (!t) return;
    t.textContent = msg;
    t.className = `tp-toast show ${kind}`;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => (t.className = "tp-toast"), 3200);
  }

  const STATE_CLASS = { idle: "idle", active: "active", testing: "active", paused: "paused", waiting_parts: "waiting", waiting_ops: "waiting", done: "done" };

  // ------------------------------------------------------------ offline queue
  function queue() {
    return read(QUEUE_KEY, []);
  }

  function setQueue(q) {
    return write(QUEUE_KEY, q);
  }

  function isNetworkError(e) {
    return !e?.status;
  }

  /**
   * Sends a write now; without signal it is kept on the phone and sent later.
   * Returns { ok, queued, data } or throws for a refusal (e.g. not allowed).
   */
  async function send(item) {
    const entry = { id: uid(), at: new Date().toISOString(), tries: 0, ...item };
    entry.body = { ...(entry.body || {}), client_event_id: entry.id, at: entry.at };
    if (navigator.onLine !== false) {
      try {
        const data = await post(entry);
        return { ok: true, data };
      } catch (e) {
        if (!isNetworkError(e)) throw e;
      }
    }
    const q = queue();
    q.push(entry);
    if (!setQueue(q)) throw new Error("Phone storage is full; could not keep this offline.");
    renderSyncBar();
    return { ok: true, queued: true };
  }

  async function post(entry) {
    if (entry.kind === "photo") {
      const blob = await (await fetch(entry.dataUrl)).blob();
      const fd = new FormData();
      fd.append("file", blob, "photo.jpg");
      const p = new URLSearchParams({ client_event_id: entry.id, at: entry.at, caption: entry.caption || "" });
      return A.fetchJson(`${api(entry.url)}?${p}`, { method: "POST", body: fd });
    }
    return A.fetchJson(api(entry.url), { method: entry.method || "POST", body: JSON.stringify(entry.body) });
  }

  async function syncQueue() {
    if (syncing || navigator.onLine === false) return;
    let q = queue();
    if (!q.length) return renderSyncBar();
    syncing = true;
    renderSyncBar();
    let sent = 0;
    const failed = [];
    try {
      for (const entry of q) {
        try {
          await post(entry);
          sent += 1;
          q = queue().filter((x) => x.id !== entry.id);
          setQueue(q);
        } catch (e) {
          if (isNetworkError(e)) break; // still offline: keep the rest in order
          // Refused by the server (e.g. job reassigned): keep it visible, do not retry forever.
          entry.tries = (entry.tries || 0) + 1;
          entry.error = e.message || String(e);
          failed.push(entry);
          q = queue().map((x) => (x.id === entry.id ? entry : x));
          setQueue(q);
        }
      }
    } finally {
      syncing = false;
    }
    if (sent) {
      toast(`Sent ${sent} saved update${sent === 1 ? "" : "s"}`, "ok");
      render();
    }
    renderSyncBar();
  }

  function renderSyncBar() {
    const bar = qs("tpSync");
    if (!bar) return;
    const q = queue();
    const refused = q.filter((x) => x.error);
    const offline = navigator.onLine === false;
    if (!q.length && !offline) {
      bar.className = "tp-sync hidden";
      return;
    }
    bar.className = `tp-sync ${refused.length ? "bad" : offline ? "off" : "wait"}`;
    const parts = [];
    if (offline) parts.push("No signal — you can keep working.");
    if (q.length - refused.length) parts.push(`${q.length - refused.length} update${q.length - refused.length === 1 ? "" : "s"} saved on this phone${syncing ? " (sending…)" : ""}.`);
    if (refused.length) parts.push(`${refused.length} could not be saved: ${esc(refused[0].error)}`);
    bar.innerHTML = `<span>${parts.join(" ")}</span>${q.length && !offline ? `<button type="button" class="tp-link" data-act="sync">Send now</button>` : ""}${refused.length ? `<button type="button" class="tp-link" data-act="dropRefused">Clear</button>` : ""}`;
  }

  // ------------------------------------------------------------ loading
  async function load(key, url) {
    try {
      const data = await A.fetchJson(api(url));
      cachePut(key, data);
      return { data, fresh: true };
    } catch (e) {
      if (!isNetworkError(e)) throw e;
      const c = cacheGet(key);
      if (c) return { data: c.data, fresh: false, cachedAt: c.at };
      throw new Error("No signal and nothing saved on this phone yet.");
    }
  }

  function staleNote(r) {
    return r.fresh ? "" : `<div class="tp-stale">Showing what was saved at ${esc(dayTime(r.cachedAt))} — no signal.</div>`;
  }

  function pendingFor(woId) {
    return queue().filter((x) => Number(x.woId) === Number(woId) && !x.error);
  }

  // ------------------------------------------------------------ navigation
  function go(name, params = {}) {
    view = { name, ...params };
    const url = new URL(window.location.href);
    ["wo", "asset", "tab"].forEach((k) => url.searchParams.delete(k));
    if (name === "job") url.searchParams.set("wo", params.id);
    if (name === "asset") url.searchParams.set("asset", params.code);
    if (["shift", "week"].includes(name)) url.searchParams.set("tab", name);
    history.replaceState(null, "", url);
    render();
    window.scrollTo(0, 0);
  }

  function setTabs() {
    document.querySelectorAll(".tp-tab").forEach((b) => b.classList.toggle("on", b.dataset.go === view.name || (view.name === "job" && b.dataset.go === "today")));
  }

  async function render() {
    setTabs();
    clearInterval(tick);
    const main = qs("tpMain");
    if (!main) return;
    if (!main.innerHTML.trim()) main.innerHTML = `<div class="tp-empty">Loading…</div>`;
    try {
      if (view.name === "today") await renderToday(main);
      else if (view.name === "job") await renderJob(main);
      else if (view.name === "shift") await renderShift(main);
      else if (view.name === "week") await renderWeek(main);
      else if (view.name === "asset") await renderAsset(main);
      else if (view.name === "scan") renderScan(main);
    } catch (e) {
      main.innerHTML = `<div class="tp-card tp-error">${esc(e.message || e)}</div><button class="tp-btn" data-go="today">Back to Today</button>`;
    }
    renderSyncBar();
  }

  // ------------------------------------------------------------ Today
  function jobCard(c) {
    const st = STATE_CLASS[c.my_state] || "idle";
    const pend = pendingFor(c.id).length;
    const partsLine = c.parts?.short ? `${c.parts.short} part${c.parts.short === 1 ? "" : "s"} short` : c.parts?.open_requests ? `${c.parts.open_requests} part request${c.parts.open_requests === 1 ? "" : "s"} open` : "";
    return `
      <button type="button" class="tp-job ${c.critical ? "crit" : ""}" data-open="${c.id}">
        <div class="tp-job-top">
          <span class="tp-asset">${esc(c.asset_code)}</span>
          <span class="tp-chip ${st}">${esc(c.my_state_label)}</span>
        </div>
        <div class="tp-job-line">${esc(c.job)}</div>
        <div class="tp-job-meta">
          <span>WO #${esc(c.id)}</span>
          ${c.role === "helper" ? `<span>Helper</span>` : ""}
          ${c.due_date ? `<span>Due ${esc(String(c.due_date).slice(0, 10))}</span>` : ""}
          ${c.my_hours_today ? `<span>${esc(hrs(c.my_hours_today))} today</span>` : ""}
          ${partsLine ? `<span class="warn">${esc(partsLine)}</span>` : ""}
          ${pend ? `<span class="warn">${pend} not sent</span>` : ""}
        </div>
      </button>`;
  }

  function section(title, list, note = "") {
    if (!list.length) return "";
    return `<h3 class="tp-h">${esc(title)} <span class="tp-count">${list.length}</span></h3>${note}${list.map(jobCard).join("")}`;
  }

  async function renderToday(main) {
    const r = await load("today", "/today");
    const d = r.data;
    const cur = d.current;
    const g = d.groups || { urgent: [], planned: [], waiting: [], completed: [] };
    const nothing = !g.urgent.length && !g.planned.length && !g.waiting.length;
    main.innerHTML = `
      ${staleNote(r)}
      <div class="tp-hello">
        <div>
          <div class="tp-hello-name">${esc(d.user?.name || A.getSessionUser())}</div>
          <div class="tp-muted">${d.shift ? `Shift since ${esc(time(d.shift.started_at))} · ${esc(hrs(d.hours_today))} on jobs today` : "Shift not started"}</div>
        </div>
        ${d.shift ? `<button type="button" class="tp-btn sm" data-go="shift">Shift report</button>` : `<button type="button" class="tp-btn sm primary" data-act="shiftStart">Start shift</button>`}
      </div>
      ${cur ? `
        <div class="tp-current" data-open="${cur.id}">
          <div class="tp-muted small">${cur.state === "testing" ? "Testing now" : "Working on now"}</div>
          <div class="tp-current-asset">${esc(cur.asset_code)} <span class="tp-muted">WO #${esc(cur.id)}</span></div>
          <div>${esc(cur.job)}</div>
          <div class="tp-timer" data-since="${esc(cur.since)}">${esc(since(cur.since))}</div>
          <div class="tp-row">
            <button type="button" class="tp-btn" data-quick="pause" data-wo="${cur.id}">Pause</button>
            <button type="button" class="tp-btn primary" data-open="${cur.id}">Open job</button>
          </div>
        </div>` : ""}
      ${(d.notifications || []).map((n) => `<div class="tp-note ${esc(n.kind)}" data-open="${n.wo_id}">${esc(n.text)}</div>`).join("")}
      <div class="tp-stats">
        <div><b>${d.counts.urgent}</b><span>Urgent</span></div>
        <div><b>${d.counts.planned}</b><span>Planned</span></div>
        <div><b>${d.counts.waiting}</b><span>Waiting</span></div>
        <div><b>${d.counts.completed}</b><span>Done today</span></div>
      </div>
      ${section("Urgent", g.urgent)}
      ${section("Planned", g.planned)}
      ${section("Waiting", g.waiting)}
      ${nothing ? `<div class="tp-card tp-empty">No open jobs for you. Your foreman assigns jobs on the Work Orders board.</div>` : ""}
      ${section("Completed today", g.completed)}
      <div class="tp-card">
        <div class="tp-row">
          <input id="tpWoNo" type="number" inputmode="numeric" min="1" placeholder="Open WO #" class="tp-input grow" />
          <button type="button" class="tp-btn" data-act="openWoNo">Open</button>
        </div>
      </div>`;
    startTimers();
  }

  function startTimers() {
    clearInterval(tick);
    tick = setInterval(() => {
      document.querySelectorAll("[data-since]").forEach((el) => (el.textContent = since(el.dataset.since)));
    }, 30000);
  }

  // ------------------------------------------------------------ Job
  const ACTIONS = {
    idle: [["start", "Start job", "primary big"]],
    active: [["pause", "Pause", ""], ["waiting_parts", "Waiting for parts", "warn"], ["waiting_ops", "Waiting for operations", "warn"], ["testing", "Testing", ""], ["complete", "Complete job", "ok big"]],
    testing: [["resume", "Back to work", ""], ["pause", "Pause", ""], ["complete", "Complete job", "ok big"]],
    paused: [["resume", "Resume", "primary big"], ["waiting_parts", "Waiting for parts", "warn"], ["waiting_ops", "Waiting for operations", "warn"], ["complete", "Complete job", "ok"]],
    waiting_parts: [["resume", "Resume", "primary big"], ["waiting_ops", "Waiting for operations", "warn"], ["complete", "Complete job", "ok"]],
    waiting_ops: [["resume", "Resume", "primary big"], ["waiting_parts", "Waiting for parts", "warn"], ["complete", "Complete job", "ok"]],
    done: [],
  };

  async function renderJob(main) {
    const id = Number(view.id);
    const r = await load(`wo:${id}`, `/workorders/${id}`);
    const d = r.data;
    const tab = view.tab || "work";
    const me = d.me;
    const finished = ["completed", "approved", "closed"].includes(String(d.wo.status).toLowerCase());
    const pend = pendingFor(id);
    let state = me.state;
    // Show the last action taken offline as the current state.
    const lastPending = [...pend].reverse().find((x) => x.kind === "action");
    if (lastPending) state = lastPending.state || state;
    const actions = finished ? [] : (ACTIONS[state] || []).filter(([a]) => !(a === "start" && String(d.wo.status).toLowerCase() === "open"));
    const running = ["active", "testing"].includes(state);
    const team = d.team || {};
    main.innerHTML = `
      ${staleNote(r)}
      <div class="tp-jobhead">
        <button type="button" class="tp-back" data-go="today">‹ Today</button>
        <div class="tp-jobhead-main">
          <div class="tp-asset lg">${esc(d.asset.code)} <span class="tp-muted">${esc(d.asset.name || "")}</span></div>
          <div class="tp-job-line">${esc(d.wo.job)}</div>
          <div class="tp-job-meta">
            <span>WO #${esc(d.wo.id)}</span>
            <span>${esc(String(d.wo.status).replace(/_/g, " "))}</span>
            ${Number(d.asset.meter?.hours) > 0 ? `<span>${esc(Number(d.asset.meter.hours).toFixed(0))} h on meter</span>` : ""}
            ${me.role !== "lead" ? `<span>${me.role === "helper" ? "You are helping" : "Viewing as foreman"}</span>` : ""}
          </div>
        </div>
      </div>
      <div class="tp-state ${STATE_CLASS[state] || "idle"}">
        <div>
          <div class="tp-state-label">${esc(finished ? "Job completed" : lastPending ? lastPending.label : me.state_label)}</div>
          <div class="tp-muted small">
            You: ${esc(hrs(d.time.my_hours))} · Everyone: ${esc(hrs(d.time.total_hours))}
            ${running && d.time.running_since ? ` · running <span data-since="${esc(d.time.running_since)}">${esc(since(d.time.running_since))}</span>` : ""}
          </div>
        </div>
        ${pend.length ? `<span class="tp-chip waiting">${pend.length} not sent</span>` : ""}
      </div>
      ${actions.length ? `<div class="tp-actions">${actions.map(([a, label, cls]) => `<button type="button" class="tp-btn ${cls}" data-action="${a}">${esc(label)}</button>`).join("")}</div>` : ""}
      ${!finished && String(d.wo.status).toLowerCase() === "open" ? `<div class="tp-card tp-muted">This job is not assigned yet. Ask your foreman to assign it.</div>` : ""}
      <div class="tp-subtabs">
        ${[["work", "Job"], ["parts", `Parts${d.parts.summary.short || d.parts.summary.open_requests ? " •" : ""}`], ["notes", `Notes & photos${d.findings.length + d.photos.length ? ` (${d.findings.length + d.photos.length})` : ""}`], ["history", "History"], ["borris", "Ask Borris"]]
          .map(([k, label]) => `<button type="button" class="tp-subtab ${tab === k ? "on" : ""}" data-tab="${k}">${esc(label)}</button>`).join("")}
      </div>
      <div id="tpJobTab">${jobTab(tab, d)}</div>`;
    startTimers();
  }

  function jobTab(tab, d) {
    if (tab === "parts") return partsTab(d);
    if (tab === "notes") return notesTab(d);
    if (tab === "history") return historyTab(d);
    if (tab === "borris") return borrisTab(d);
    const b = d.breakdown;
    const s = d.service;
    const team = d.team || {};
    return `
      ${b ? `<div class="tp-card">
        <div class="tp-h4">Breakdown ${b.critical ? `<span class="tp-chip bad">Critical</span>` : ""}</div>
        <div>${esc([b.component, b.description].filter(Boolean).join(" — "))}</div>
        <div class="tp-muted small">${b.since ? `Down since ${esc(dayTime(b.since))}` : ""}${b.return_date ? ` · Return date ${esc(b.return_date)}` : ""}${b.parts_status ? ` · Parts: ${esc(b.parts_status)}` : ""}</div>
      </div>` : ""}
      ${s ? `<div class="tp-card">
        <div class="tp-h4">Service</div>
        <div>${esc(s.name)} · every ${esc(s.interval)} h</div>
        <div class="tp-muted small">Due at ${esc(s.next_due)} h · ${s.remaining >= 0 ? `${esc(s.remaining)} h to go` : `${esc(Math.abs(s.remaining))} h overdue`}</div>
      </div>` : ""}
      ${d.wo.job_description ? `<div class="tp-card"><div class="tp-h4">Job card</div><div class="tp-pre">${esc(d.wo.job_description)}</div></div>` : ""}
      ${d.wo.progress ? `<div class="tp-card"><div class="tp-h4">Latest progress</div><div>${esc(d.wo.progress)}</div></div>` : ""}
      ${d.wo.completion_notes ? `<div class="tp-card"><div class="tp-h4">What was done</div><div>${esc(d.wo.completion_notes)}</div></div>` : ""}
      <div class="tp-card">
        <div class="tp-h4">Team</div>
        <div>Lead: ${esc(team.lead_name || team.lead || "not assigned")}</div>
        ${(team.helpers || []).map((h) => `<div>Helper: ${esc(h.name)} <span class="tp-muted small">${esc(labelOf(team.states?.[h.username]))}</span>
          ${isLead() ? `<button type="button" class="tp-link" data-act="helperDel" data-user="${esc(h.username)}">Remove</button>` : ""}</div>`).join("")}
        ${isLead() ? `<div class="tp-row mt"><input id="tpHelper" class="tp-input grow" placeholder="Add helper (username)" /><button type="button" class="tp-btn sm" data-act="helperAdd">Add</button></div>` : ""}
      </div>`;
  }

  function labelOf(st) {
    return { idle: "not started", active: "working", testing: "testing", paused: "paused", waiting_parts: "waiting for parts", waiting_ops: "waiting for operations", done: "finished" }[st] || "not started";
  }

  function partsTab(d) {
    const p = d.parts;
    const stateChip = { issued: ["ok", "Issued"], in_stock: ["active", "In stores"], short: ["bad", "Short"] };
    return `
      ${p.planned.length ? `<div class="tp-card"><div class="tp-h4">Planned for this job</div>
        ${p.planned.map((l) => `<div class="tp-part">
          <div><b>${esc(l.part_code || "")}</b> ${esc(l.description || "")}</div>
          <div class="tp-muted small">Need ${esc(l.quantity_planned)} ${esc(l.unit_of_measure || "")} · issued ${esc(l.quantity_issued)} · in stores ${l.on_hand == null ? "?" : esc(l.on_hand)}${l.bin ? ` · bin ${esc(l.bin)}` : ""}</div>
          <span class="tp-chip ${stateChip[l.state][0]}">${stateChip[l.state][1]}</span>
        </div>`).join("")}</div>` : ""}
      ${p.issued.length ? `<div class="tp-card"><div class="tp-h4">Issued to this job</div>
        ${p.issued.map((l) => `<div class="tp-part"><div><b>${esc(l.part_code)}</b> ${esc(l.part_name)}</div><div class="tp-muted small">${esc(l.qty)} issued${l.bin ? ` · from ${esc(l.bin)}` : ""}</div></div>`).join("")}</div>` : ""}
      <div class="tp-card"><div class="tp-h4">Requests to stores</div>
        ${p.requests.length ? p.requests.map((r) => `<div class="tp-part"><div><b>${esc(r.part_code || "")}</b> ${esc(r.part_name || "")} × ${esc(r.qty)}</div>
          <div class="tp-muted small">${esc(r.requested_by || "")} · ${esc(dayTime(r.created_at))}${r.status_notes ? ` · ${esc(r.status_notes)}` : ""}</div>
          <span class="tp-chip ${String(r.status).toLowerCase() === "received" ? "ok" : String(r.status).toLowerCase() === "ordered" ? "active" : "waiting"}">${esc(r.status || "requested")}</span></div>`).join("") : `<div class="tp-muted">No requests yet.</div>`}
      </div>
      <div class="tp-card">
        <div class="tp-h4">Request a part</div>
        <p class="tp-muted small">Stores issue parts to the job. Your request goes to the stores queue.</p>
        <input id="tpPartQ" class="tp-input" placeholder="Search stores (code or name)" autocomplete="off" />
        <div id="tpPartHits"></div>
        <input id="tpPartCode" class="tp-input" placeholder="Part code (if known)" />
        <input id="tpPartName" class="tp-input" placeholder="What part is needed" />
        <div class="tp-row">
          <input id="tpPartQty" class="tp-input" type="number" inputmode="decimal" min="1" value="1" style="max-width:90px" />
          <select id="tpPartUrg" class="tp-input"><option value="normal">Normal</option><option value="urgent">Urgent — machine down</option></select>
        </div>
        <input id="tpPartNote" class="tp-input" placeholder="Note for stores (optional)" />
        <button type="button" class="tp-btn primary full" data-act="partRequest">Send to stores</button>
      </div>`;
  }

  function notesTab(d) {
    return `
      <div class="tp-card">
        <div class="tp-h4">Add a finding</div>
        <textarea id="tpFinding" class="tp-input" rows="3" placeholder="What did you find? e.g. hose chafed on the chassis clamp"></textarea>
        <div class="tp-row">
          <select id="tpFindingKind" class="tp-input"><option value="finding">Finding</option><option value="safety">Safety issue</option><option value="note">Note</option></select>
          <button type="button" class="tp-btn primary" data-act="finding">Save</button>
        </div>
        <label class="tp-btn full tp-photo-btn">📷 Add photo<input id="tpPhoto" type="file" accept="image/*" capture="environment" hidden /></label>
      </div>
      ${d.photos.length ? `<div class="tp-photos">${d.photos.map((p) => `<a href="${esc(p.url)}" target="_blank" rel="noopener"><img src="${esc(p.url)}" alt="${esc(p.caption || "Job photo")}" loading="lazy" /></a>`).join("")}</div>` : ""}
      ${d.findings.map((f) => `<div class="tp-card tp-finding ${esc(f.kind)}"><div class="tp-muted small">${esc(f.kind === "safety" ? "Safety" : f.kind === "note" ? "Note" : "Finding")} · ${esc(f.username)} · ${esc(dayTime(f.at))}</div><div>${esc(f.text)}</div></div>`).join("")}`;
  }

  function historyTab(d) {
    return `
      <div class="tp-card"><div class="tp-h4">Earlier jobs on ${esc(d.asset.code)}</div>
        ${d.history.length ? d.history.map((h) => `<div class="tp-hist"><div><b>WO #${esc(h.id)}</b> ${esc(h.job)}</div><div class="tp-muted small">${esc(dayTime(h.done_at))}${h.notes ? ` · ${esc(h.notes)}` : ""}</div></div>`).join("") : `<div class="tp-muted">No earlier jobs recorded.</div>`}
      </div>
      <div class="tp-card"><div class="tp-h4">Activity on this job</div>
        ${d.activity.length ? d.activity.map((a) => `<div class="tp-hist"><div>${esc(a.text)}</div><div class="tp-muted small">${esc(a.who)} · ${esc(dayTime(a.at))}${a.auto ? " · automatic" : ""}</div></div>`).join("") : `<div class="tp-muted">Nothing yet.</div>`}
      </div>
      <button type="button" class="tp-btn full" data-asset="${esc(d.asset.code)}">Machine view: ${esc(d.asset.code)}</button>`;
  }

  function borrisTab(d) {
    const ideas = d.breakdown
      ? ["What usually causes this failure?", "What should I check first?", "Which manual section covers this?"]
      : ["What does this service include?", "Which manual section covers this?", "What torque or fluid spec applies?"];
    return `
      <div class="tp-card">
        <div class="tp-h4">Ask Borris about this job</div>
        <p class="tp-muted small">Borris gives advice only. He cannot change the job, stock or the service plan.</p>
        <div class="tp-ideas">${ideas.map((q) => `<button type="button" class="tp-idea" data-idea="${esc(q)}">${esc(q)}</button>`).join("")}</div>
        <textarea id="tpAsk" class="tp-input" rows="2" placeholder="Your question"></textarea>
        <button type="button" class="tp-btn primary full" data-act="ask">Ask</button>
        <div id="tpAnswer"></div>
      </div>`;
  }

  // ------------------------------------------------------------ sheets
  function sheet(html) {
    const s = qs("tpSheet");
    s.innerHTML = `<div class="tp-sheet-card">${html}</div>`;
    s.classList.remove("hidden");
    return s;
  }

  function closeSheet() {
    qs("tpSheet")?.classList.add("hidden");
  }

  const ACTION_LABEL = { start: "Working", resume: "Working", pause: "Paused", waiting_parts: "Waiting for parts", waiting_ops: "Waiting for operations", testing: "Testing", complete: "Completed" };
  const ACTION_STATE = { start: "active", resume: "active", pause: "paused", waiting_parts: "waiting_parts", waiting_ops: "waiting_ops", testing: "testing", complete: "done" };

  async function doAction(woId, action, extra = {}) {
    const res = await send({ kind: "action", woId, url: `/workorders/${woId}/action`, body: { action, ...extra }, label: ACTION_LABEL[action], state: ACTION_STATE[action] });
    toast(res.queued ? `${ACTION_LABEL[action]} — saved on this phone, will send when there is signal` : `${ACTION_LABEL[action]}`, res.queued ? "" : "ok");
    return res;
  }

  function askAction(woId, action, d) {
    if (action === "complete") {
      const lead = d?.me?.role !== "helper";
      sheet(`
        <h3>${lead ? "Complete the job" : "Finish your part"}</h3>
        ${lead ? `<label class="tp-label">What was done <b>(required)</b></label>
        <textarea id="tpDone" class="tp-input" rows="4" placeholder="e.g. Replaced hydraulic hose and clamp, tested under load, no leaks"></textarea>
        <label class="tp-label">Labour hours (leave empty to use the job timer: ${esc(hrs(d?.time?.total_hours || 0))})</label>
        <input id="tpDoneHours" class="tp-input" type="number" inputmode="decimal" min="0" step="0.25" placeholder="${esc(Number(d?.time?.total_hours || 0).toFixed(2))}" />` : `<p class="tp-muted">The lead technician completes the work order. This stops your time on it.</p>
        <textarea id="tpDone" class="tp-input" rows="2" placeholder="What you did (optional)"></textarea>`}
        <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">Cancel</button><button type="button" class="tp-btn ok grow" data-act="confirmComplete" data-wo="${woId}" data-lead="${lead ? 1 : 0}">${lead ? "Complete job" : "Finish my part"}</button></div>`);
      return;
    }
    if (action === "waiting_parts" || action === "waiting_ops" || action === "pause") {
      const title = { waiting_parts: "Waiting for parts", waiting_ops: "Waiting for operations", pause: "Pause the job" }[action];
      const hint = { waiting_parts: "Which part? e.g. hose from stores, ordered from Maputo", waiting_ops: "e.g. machine needed on the bench, waiting for the operator", pause: "Why? (optional) e.g. lunch, called to another job" }[action];
      sheet(`
        <h3>${esc(title)}</h3>
        <textarea id="tpWhy" class="tp-input" rows="2" placeholder="${esc(hint)}"></textarea>
        <p class="tp-muted small">Time stops counting as labour until you resume.</p>
        <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">Cancel</button><button type="button" class="tp-btn primary grow" data-act="confirmWhy" data-wo="${woId}" data-action="${action}">${esc(title)}</button></div>`);
      return;
    }
    return doAction(woId, action).then(render).catch((e) => toast(e.message || e, "bad"));
  }

  // ------------------------------------------------------------ Shift
  const SHIFT_FIELDS = [
    ["findings", "Findings", "Anything you found on machines this shift"],
    ["unplanned", "Unplanned work", "Work done that was not on a job card"],
    ["safety", "Safety", "Near misses, hazards, PPE issues"],
    ["outstanding", "Outstanding work", "What is not finished and why"],
    ["handover", "Handover", "What the next shift must know"],
    ["next_shift", "For the next shift", "Jobs to pick up first"],
  ];

  async function renderShift(main) {
    const r = await load("shift", "/shift");
    const s = r.data.shift;
    if (!s) {
      main.innerHTML = `${staleNote(r)}
        <h2 class="tp-title">Shift report</h2>
        <div class="tp-card"><p>Your shift has not started. It starts when you tap Start shift or start your first job, and ends when you submit the report — a night shift over midnight is one report.</p>
        <button type="button" class="tp-btn primary big full" data-act="shiftStart">Start shift</button></div>
        ${isLead() ? `<button type="button" class="tp-btn full" data-act="teamShifts">Team shift reports</button><div id="tpTeam"></div>` : ""}`;
      return;
    }
    const labour = s.timeline.filter((t) => t.labour);
    main.innerHTML = `${staleNote(r)}
      <h2 class="tp-title">Shift report</h2>
      <div class="tp-card">
        <div class="tp-row between"><div>Started ${esc(dayTime(s.started_at))}</div><div class="tp-big">${esc(hrs(s.hours))}</div></div>
        <div class="tp-muted small">on jobs · ${labour.length} work block${labour.length === 1 ? "" : "s"} · ${s.findings_logged.length} finding${s.findings_logged.length === 1 ? "" : "s"} logged</div>
      </div>
      <div class="tp-card"><div class="tp-h4">What you did</div>
        ${s.timeline.length ? s.timeline.map((t) => `<div class="tp-tl ${STATE_CLASS[t.state] || ""}">
          <div class="tp-tl-time">${esc(time(t.start))}–${t.open ? "now" : esc(time(t.end))}</div>
          <div><b>${esc(t.asset_code || "")}</b> ${esc(t.job || "")} <span class="tp-muted small">WO #${esc(t.work_order_id)}</span><div class="tp-muted small">${esc(t.state_label)} · ${esc(hrs(t.hours))}</div></div>
        </div>`).join("") : `<div class="tp-muted">No job time yet this shift.</div>`}
      </div>
      ${s.findings_logged.length ? `<div class="tp-card"><div class="tp-h4">Findings logged on jobs</div>${s.findings_logged.map((f) => `<div class="tp-hist">${esc(f.text)} <span class="tp-muted small">${f.work_order_id ? `WO #${esc(f.work_order_id)}` : ""} ${esc(time(f.at))}</span></div>`).join("")}</div>` : ""}
      <div class="tp-card">
        ${SHIFT_FIELDS.map(([k, label, hint]) => `<label class="tp-label">${esc(label)}</label><textarea class="tp-input" rows="2" data-shift="${k}" placeholder="${esc(hint)}">${esc(s[k] || "")}</textarea>`).join("")}
        <div class="tp-muted small">Saved as you type.</div>
      </div>
      <button type="button" class="tp-btn ok big full" data-act="shiftSubmit">Submit shift report</button>
      ${isLead() ? `<button type="button" class="tp-btn full" data-act="teamShifts">Team shift reports</button><div id="tpTeam"></div>` : ""}`;
  }

  let shiftSave = null;
  function queueShiftSave() {
    clearTimeout(shiftSave);
    shiftSave = setTimeout(async () => {
      const body = {};
      document.querySelectorAll("[data-shift]").forEach((el) => (body[el.dataset.shift] = el.value));
      try {
        await A.fetchJson(api("/shift"), { method: "PUT", body: JSON.stringify(body) });
      } catch (e) {
        if (isNetworkError(e)) {
          // Keep the text on the phone; it goes with the submit.
          write("ironlog-tech-shift-draft", body);
        }
      }
    }, 900);
  }

  async function renderTeam() {
    const host = qs("tpTeam");
    if (!host) return;
    host.innerHTML = `<div class="tp-muted">Loading…</div>`;
    const d = await A.fetchJson(api("/shifts?days=7"));
    host.innerHTML = d.rows.length ? d.rows.map((s) => `<details class="tp-card">
      <summary><b>${esc(s.name)}</b> · ${esc(dayTime(s.started_at))} → ${esc(time(s.ended_at))} · ${esc(hrs(s.hours))}</summary>
      ${s.timeline.filter((t) => t.labour).map((t) => `<div class="tp-muted small">${esc(time(t.start))}–${esc(time(t.end))} ${esc(t.asset_code || "")} WO #${esc(t.work_order_id)} ${esc(t.job || "")} (${esc(hrs(t.hours))})</div>`).join("")}
      ${SHIFT_FIELDS.filter(([k]) => s[k]).map(([k, label]) => `<div class="mt"><b>${esc(label)}:</b> ${esc(s[k])}</div>`).join("")}
    </details>`).join("") : `<div class="tp-muted">No submitted reports in the last 7 days.</div>`;
  }

  // ------------------------------------------------------------ Week
  async function renderWeek(main) {
    const r = await load("week", "/week");
    const d = r.data;
    const max = Math.max(8, ...d.days.map((x) => x.hours));
    main.innerHTML = `${staleNote(r)}
      <h2 class="tp-title">My week</h2>
      <div class="tp-stats">
        <div><b>${esc(Number(d.hours).toFixed(1))}</b><span>Hours on jobs</span></div>
        <div><b>${d.jobs_completed}</b><span>Completed</span></div>
        <div><b>${d.jobs_open}</b><span>Open</span></div>
        <div><b>${d.jobs_waiting}</b><span>Waiting</span></div>
      </div>
      <div class="tp-card">
        ${d.days.map((x) => `<div class="tp-bar"><span class="tp-bar-day">${esc(new Date(`${x.day}T12:00:00`).toLocaleDateString([], { weekday: "short", day: "2-digit" }))}</span><span class="tp-bar-track"><span style="width:${Math.round((x.hours / max) * 100)}%"></span></span><span class="tp-bar-val">${esc(hrs(x.hours))}</span></div>`).join("")}
      </div>
      <div class="tp-card tp-muted small">${d.shifts_submitted} shift report${d.shifts_submitted === 1 ? "" : "s"} submitted this week${d.shift_open ? " · a shift is open now" : ""}.</div>`;
  }

  // ------------------------------------------------------------ Machine (QR)
  function renderScan(main) {
    main.innerHTML = `
      <h2 class="tp-title">Machine</h2>
      <div class="tp-card">
        <p class="tp-muted">Scan the machine's QR label with the phone camera, or type the fleet number.</p>
        <div class="tp-row"><input id="tpAssetCode" class="tp-input grow" placeholder="e.g. T01AM" autocapitalize="characters" /><button type="button" class="tp-btn primary" data-act="openAsset">Open</button></div>
      </div>`;
  }

  async function renderAsset(main) {
    const code = String(view.code || "").toUpperCase();
    const r = await load(`asset:${code}`, `/assets/${encodeURIComponent(code)}`);
    const d = r.data;
    const a = d.asset;
    main.innerHTML = `${staleNote(r)}
      <div class="tp-jobhead"><button type="button" class="tp-back" data-go="scan">‹ Machines</button>
        <div class="tp-jobhead-main"><div class="tp-asset lg">${esc(a.asset_code)} <span class="tp-muted">${esc(a.asset_name || "")}</span></div>
        <div class="tp-job-meta">${Number(a.meter?.hours) > 0 ? `<span>${esc(Number(a.meter.hours).toFixed(0))} h on meter</span>` : ""}${d.breakdown ? `<span class="bad">Broken down</span>` : `<span>Running</span>`}</div></div></div>
      ${d.breakdown ? `<div class="tp-card"><div class="tp-h4">Open breakdown ${d.breakdown.critical ? `<span class="tp-chip bad">Critical</span>` : ""}</div><div>${esc([d.breakdown.component, d.breakdown.description].filter(Boolean).join(" — "))}</div><div class="tp-muted small">Since ${esc(dayTime(d.breakdown.start_at || d.breakdown.breakdown_date))}${d.breakdown.ets_repair_date ? ` · return date ${esc(d.breakdown.ets_repair_date)}` : ""}</div></div>` : ""}
      <div class="tp-card"><div class="tp-h4">Open jobs</div>
        ${d.active.length ? d.active.map((w) => `<button type="button" class="tp-job" data-open="${w.id}"><div class="tp-job-top"><span>WO #${esc(w.id)}</span><span class="tp-chip ${w.mine ? "active" : "idle"}">${w.mine ? "Yours" : esc(w.lead || "Unassigned")}</span></div><div class="tp-job-line">${esc(w.job)}</div></button>`).join("") : `<div class="tp-muted">No open jobs on this machine.</div>`}
      </div>
      ${d.next_service ? `<div class="tp-card"><div class="tp-h4">Next service</div><div>${esc(d.next_service.service)} at ${esc(d.next_service.due_at)} h</div><div class="tp-muted small">${d.next_service.remaining >= 0 ? `${esc(d.next_service.remaining)} h to go` : `${esc(Math.abs(d.next_service.remaining))} h overdue`}</div></div>` : ""}
      <div class="tp-card"><div class="tp-h4">Recent work</div>${d.history.length ? d.history.map((h) => `<div class="tp-hist"><div><b>WO #${esc(h.id)}</b> ${esc(h.job)}</div><div class="tp-muted small">${esc(dayTime(h.done_at))}${h.notes ? ` · ${esc(h.notes)}` : ""}</div></div>`).join("") : `<div class="tp-muted">No completed jobs yet.</div>`}</div>
      ${d.findings.length ? `<div class="tp-card"><div class="tp-h4">Findings</div>${d.findings.map((f) => `<div class="tp-hist">${esc(f.text)} <span class="tp-muted small">${esc(f.username)} · ${esc(dayTime(f.at))}</span></div>`).join("")}</div>` : ""}
      <div class="tp-card">
        <div class="tp-h4">Found a problem?</div>
        <textarea id="tpAssetFinding" class="tp-input" rows="2" placeholder="What did you find?"></textarea>
        <div class="tp-row"><button type="button" class="tp-btn grow" data-act="assetFinding" data-asset-id="${a.id}">Save finding</button><button type="button" class="tp-btn warn grow" data-act="breakdownForm" data-code="${esc(a.asset_code)}">Report breakdown</button></div>
      </div>`;
  }

  // ------------------------------------------------------------ photo
  function compress(file) {
    return new Promise((resolve, reject) => {
      const img = new Image();
      const url = URL.createObjectURL(file);
      img.onload = () => {
        const scale = Math.min(1, 1600 / Math.max(img.width, img.height));
        const c = document.createElement("canvas");
        c.width = Math.round(img.width * scale);
        c.height = Math.round(img.height * scale);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
        URL.revokeObjectURL(url);
        resolve(c.toDataURL("image/jpeg", 0.78));
      };
      img.onerror = () => reject(new Error("Could not read that photo"));
      img.src = url;
    });
  }

  // ------------------------------------------------------------ events
  async function onClick(e) {
    const t = e.target.closest("[data-go],[data-open],[data-action],[data-act],[data-tab],[data-asset],[data-quick],[data-idea]");
    if (!t) return;
    const cached = view.name === "job" ? cacheGet(`wo:${view.id}`)?.data : null;
    try {
      if (t.dataset.go) return go(t.dataset.go);
      if (t.dataset.open) return go("job", { id: Number(t.dataset.open) });
      if (t.dataset.asset) return go("asset", { code: t.dataset.asset });
      if (t.dataset.tab) {
        view.tab = t.dataset.tab;
        document.querySelectorAll(".tp-subtab").forEach((b) => b.classList.toggle("on", b.dataset.tab === view.tab));
        if (cached) qs("tpJobTab").innerHTML = jobTab(view.tab, cached);
        return;
      }
      if (t.dataset.idea) {
        qs("tpAsk").value = t.dataset.idea;
        return;
      }
      if (t.dataset.quick) {
        await doAction(Number(t.dataset.wo), t.dataset.quick);
        return render();
      }
      if (t.dataset.action && !t.dataset.act) return askAction(view.id, t.dataset.action, cached);
      const act = t.dataset.act;
      if (act === "closeSheet") return closeSheet();
      if (act === "sync") return syncQueue();
      if (act === "dropRefused") {
        setQueue(queue().filter((x) => !x.error));
        return renderSyncBar();
      }
      if (act === "openWoNo") {
        const n = Number(qs("tpWoNo")?.value || 0);
        if (n > 0) go("job", { id: n });
        return;
      }
      if (act === "openAsset") {
        const c = String(qs("tpAssetCode")?.value || "").trim();
        if (c) go("asset", { code: c.toUpperCase() });
        return;
      }
      if (act === "confirmComplete") {
        const notes = String(qs("tpDone")?.value || "").trim();
        const lead = t.dataset.lead === "1";
        if (lead && !notes) return toast("Say what was done first.", "bad");
        const hours = Number(qs("tpDoneHours")?.value || 0);
        const extra = lead ? { completion_notes: notes, ...(hours > 0 ? { labor_hours: hours } : {}) } : { note: notes || null };
        t.disabled = true;
        await doAction(Number(t.dataset.wo), "complete", extra);
        closeSheet();
        return render();
      }
      if (act === "confirmWhy") {
        const note = String(qs("tpWhy")?.value || "").trim();
        t.disabled = true;
        await doAction(Number(t.dataset.wo), t.dataset.action, note ? { note } : {});
        closeSheet();
        return render();
      }
      if (act === "finding") {
        const text = String(qs("tpFinding")?.value || "").trim();
        if (!text) return toast("Write what you found first.", "bad");
        const res = await send({ kind: "finding", woId: view.id, url: "/findings", body: { work_order_id: view.id, text, kind: qs("tpFindingKind").value } });
        toast(res.queued ? "Saved on this phone — will send later" : "Finding saved", res.queued ? "" : "ok");
        return render();
      }
      if (act === "assetFinding") {
        const text = String(qs("tpAssetFinding")?.value || "").trim();
        if (!text) return toast("Write what you found first.", "bad");
        const res = await send({ kind: "finding", url: "/findings", body: { asset_id: Number(t.dataset.assetId), text } });
        toast(res.queued ? "Saved on this phone — will send later" : "Finding saved", res.queued ? "" : "ok");
        return render();
      }
      if (act === "breakdownForm") {
        sheet(`<h3>Report breakdown on ${esc(t.dataset.code)}</h3>
          <label class="tp-label">What failed</label><input id="tpBdComp" class="tp-input" placeholder="e.g. Hydraulics, engine, tyre" />
          <label class="tp-label">What happened</label><textarea id="tpBdDesc" class="tp-input" rows="3"></textarea>
          <label class="tp-check"><input id="tpBdCrit" type="checkbox" /> Machine cannot work (critical)</label>
          <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">Cancel</button><button type="button" class="tp-btn warn grow" data-act="breakdownSend" data-code="${esc(t.dataset.code)}">Report breakdown</button></div>`);
        return;
      }
      if (act === "breakdownSend") {
        const description = String(qs("tpBdDesc")?.value || "").trim();
        if (!description) return toast("Describe what happened.", "bad");
        t.disabled = true;
        const res = await send({ kind: "breakdown", url: "/breakdowns", body: { asset_code: t.dataset.code, component: qs("tpBdComp").value, description, critical: qs("tpBdCrit").checked } });
        closeSheet();
        toast(res.queued ? "Breakdown saved on this phone — will send later" : `Breakdown reported${res.data?.work_order_id ? ` — WO #${res.data.work_order_id}` : ""}`, res.queued ? "" : "ok");
        return render();
      }
      if (act === "partRequest") {
        const body = {
          part_code: qs("tpPartCode").value.trim(),
          part_name: qs("tpPartName").value.trim(),
          qty: Number(qs("tpPartQty").value || 1),
          urgency: qs("tpPartUrg").value,
          notes: qs("tpPartNote").value.trim(),
        };
        if (!body.part_code && !body.part_name) return toast("Say which part is needed.", "bad");
        t.disabled = true;
        const res = await send({ kind: "parts", woId: view.id, url: `/workorders/${view.id}/parts-request`, body });
        toast(res.queued ? "Request saved on this phone — will send later" : "Sent to stores", res.queued ? "" : "ok");
        return render();
      }
      if (act === "pickPart") {
        qs("tpPartCode").value = t.dataset.code;
        qs("tpPartName").value = t.dataset.name;
        qs("tpPartHits").innerHTML = "";
        return;
      }
      if (act === "helperAdd") {
        const u = String(qs("tpHelper")?.value || "").trim();
        if (!u) return;
        await A.fetchJson(api(`/workorders/${view.id}/helpers`), { method: "POST", body: JSON.stringify({ username: u }) });
        toast("Helper added", "ok");
        return render();
      }
      if (act === "helperDel") {
        await A.fetchJson(api(`/workorders/${view.id}/helpers/${encodeURIComponent(t.dataset.user)}`), { method: "DELETE" });
        return render();
      }
      if (act === "shiftStart") {
        const res = await send({ kind: "shift", url: "/shift/start", body: {} });
        toast(res.queued ? "Shift start saved on this phone" : "Shift started", "ok");
        return render();
      }
      if (act === "shiftSubmit") {
        const s = cacheGet("shift")?.data?.shift;
        sheet(`<h3>Submit shift report</h3>
          <p>${esc(hrs(s?.hours || 0))} on jobs. Jobs still running will be paused, and your job time goes to the mechanics timesheet.</p>
          <p class="tp-muted small">After submitting, the report is locked and your foreman can see it.</p>
          <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">Cancel</button><button type="button" class="tp-btn ok grow" data-act="shiftSubmitGo">Submit</button></div>`);
        return;
      }
      if (act === "shiftSubmitGo") {
        const body = { ...read("ironlog-tech-shift-draft", {}) };
        document.querySelectorAll("[data-shift]").forEach((el) => (body[el.dataset.shift] = el.value));
        t.disabled = true;
        const res = await send({ kind: "shift", url: "/shift/submit", body });
        try { localStorage.removeItem("ironlog-tech-shift-draft"); } catch { /* ignore */ }
        closeSheet();
        toast(res.queued ? "Report saved on this phone — will send when there is signal" : `Shift submitted · ${hrs(res.data?.hours || 0)}`, "ok");
        return go("today");
      }
      if (act === "teamShifts") return renderTeam();
      if (act === "ask") {
        const q = String(qs("tpAsk")?.value || "").trim();
        if (!q) return;
        const d = cached;
        const out = qs("tpAnswer");
        out.innerHTML = `<div class="tp-muted">Borris is thinking…</div>`;
        t.disabled = true;
        const context = [
          `Work order #${d.wo.id} on ${d.asset.code} ${d.asset.name || ""} (${d.asset.category || ""}).`,
          `Job: ${d.wo.job}.`,
          d.breakdown ? `Breakdown: ${[d.breakdown.component, d.breakdown.description].filter(Boolean).join(" - ")}.` : "",
          d.service ? `Service: ${d.service.name} every ${d.service.interval} h.` : "",
          Number(d.asset.meter?.hours) > 0 ? `Meter ${Number(d.asset.meter.hours).toFixed(0)} h.` : "",
          d.findings.length ? `Findings: ${d.findings.slice(0, 5).map((f) => f.text).join("; ")}.` : "",
          d.history.length ? `Earlier jobs: ${d.history.slice(0, 4).map((h) => h.job).join("; ")}.` : "",
          "Answer for a workshop technician: short, practical steps. Advice only.",
        ].filter(Boolean).join(" ");
        try {
          const ans = await A.fetchJson(`${A.API}/api/ironmind/ask`, { method: "POST", body: JSON.stringify({ question: q, asset_code: d.asset.code, context_notes: context }) });
          const sources = (ans.sources || []).slice(0, 4).map((s) => esc(s.title || s.document_title || s.name || "")).filter(Boolean);
          out.innerHTML = `<div class="tp-answer">${esc(ans.short_answer || ans.answer || "No answer.")}</div>${sources.length ? `<div class="tp-muted small">From: ${sources.join(", ")}</div>` : ""}`;
        } catch (err) {
          out.innerHTML = `<div class="tp-error">${esc(isNetworkError(err) ? "Borris needs signal." : err.message || err)}</div>`;
        } finally {
          t.disabled = false;
        }
      }
    } catch (err) {
      t.disabled = false;
      toast(err.message || String(err), "bad");
    }
  }

  let partTimer = null;
  function onInput(e) {
    if (e.target.dataset.shift) return queueShiftSave();
    if (e.target.id === "tpPartQ") {
      clearTimeout(partTimer);
      const q = e.target.value.trim();
      partTimer = setTimeout(async () => {
        const host = qs("tpPartHits");
        if (!host) return;
        if (q.length < 2) return (host.innerHTML = "");
        try {
          const d = await A.fetchJson(api(`/parts/search?q=${encodeURIComponent(q)}`));
          host.innerHTML = d.rows.map((p) => `<button type="button" class="tp-hit" data-act="pickPart" data-code="${esc(p.part_code)}" data-name="${esc(p.part_name)}"><b>${esc(p.part_code)}</b> ${esc(p.part_name)} <span class="tp-muted small">${esc(p.on_hand)} in stores${p.bin ? ` · ${esc(p.bin)}` : ""}</span></button>`).join("") || `<div class="tp-muted small">Nothing in stores matches. Type the part below.</div>`;
        } catch {
          host.innerHTML = `<div class="tp-muted small">Search needs signal. Type the part below.</div>`;
        }
      }, 300);
    }
  }

  async function onChange(e) {
    if (e.target.id !== "tpPhoto" || !e.target.files?.[0]) return;
    try {
      const dataUrl = await compress(e.target.files[0]);
      const entry = { kind: "photo", woId: view.id, url: `/workorders/${view.id}/photos`, dataUrl };
      const res = await send(entry);
      toast(res.queued ? "Photo kept on this phone — will send later" : "Photo added", res.queued ? "" : "ok");
      render();
    } catch (err) {
      toast(err.message || String(err), "bad");
    }
  }

  // ------------------------------------------------------------ start
  function registerWorker() {
    if (!("serviceWorker" in navigator)) return;
    navigator.serviceWorker.register("./tech-sw.js", { scope: "./technician-terminal.html" }).catch(() => {});
  }

  function start() {
    root = qs("tpApp");
    if (!root || root.dataset.ready) return render();
    root.dataset.ready = "1";
    root.addEventListener("click", onClick);
    root.addEventListener("input", onInput);
    root.addEventListener("change", onChange);
    qs("tpSheet")?.addEventListener("click", (e) => {
      if (e.target.id === "tpSheet") closeSheet();
    });
    window.addEventListener("online", () => {
      renderSyncBar();
      syncQueue();
    });
    window.addEventListener("offline", renderSyncBar);
    setInterval(syncQueue, 60000);
    const p = new URLSearchParams(window.location.search);
    if (p.get("wo")) view = { name: "job", id: Number(p.get("wo")) };
    else if (p.get("asset")) view = { name: "asset", code: p.get("asset").toUpperCase() };
    else if (["shift", "week"].includes(p.get("tab"))) view = { name: p.get("tab") };
    registerWorker();
    render();
    syncQueue();
  }

  window.TechPortal = { start, syncQueue };
})();

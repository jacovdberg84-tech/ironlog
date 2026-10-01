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

  const LANG_KEY = "ironlog_prestart_lang"; // shared with the pre-start pages
  const PT = window.TechPortalPT || {};

  function lang() {
    try {
      const saved = localStorage.getItem(LANG_KEY);
      if (saved === "en" || saved === "pt") return saved;
    } catch { /* ignore */ }
    return String(navigator.language || "").toLowerCase().startsWith("pt") ? "pt" : "en";
  }

  /** Screen text in the chosen language; {name} placeholders are filled from vars. */
  function T(text, vars = {}) {
    const s = (lang() === "pt" && PT[text]) || text;
    return s.replace(/\{(\w+)\}/g, (_, k) => (vars[k] == null ? "" : String(vars[k])));
  }

  /** Singular / plural: N(n, "{n} job", "{n} jobs"). */
  function N(n, one, many) {
    return T(Number(n) === 1 ? one : many, { n });
  }

  const locale = () => (lang() === "pt" ? "pt-PT" : "en-GB");

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
    return d.toLocaleTimeString(locale(), { hour: "2-digit", minute: "2-digit", hour12: false });
  }

  function dayTime(iso) {
    if (!iso) return "";
    const d = new Date(String(iso).includes("T") ? iso : `${String(iso).replace(" ", "T")}Z`);
    if (Number.isNaN(d.getTime())) return String(iso);
    return `${d.toLocaleDateString(locale(), { day: "2-digit", month: "short" })} ${time(iso)}`;
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

  const STATE_TEXT = {
    idle: "Not started",
    active: "Working",
    paused: "Paused",
    waiting_parts: "Waiting for parts",
    waiting_ops: "Waiting for operations",
    testing: "Testing",
    done: "Completed",
  };
  const stateText = (st) => T(STATE_TEXT[st] || STATE_TEXT.idle);
  const WO_STATUS = { open: "Open", assigned: "Assigned", in_progress: "In progress", completed: "Completed", approved: "Approved", closed: "Closed" };
  const woStatus = (st) => T(WO_STATUS[String(st || "").toLowerCase()] || String(st || "").replace(/_/g, " "));

  const OTHER_REASONS = {
    housekeeping: "Workshop housekeeping",
    training: "Training / toolbox talk",
    safety_meeting: "Safety meeting",
    tools: "Tool and workshop repairs",
    travel: "Travel",
    waiting: "Waiting for work",
    other: "Other",
  };
  const reasonText = (k) => T(OTHER_REASONS[k] || OTHER_REASONS.other);

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
    if (!setQueue(q)) throw new Error(T("Phone storage is full; could not keep this offline."));
    renderSyncBar();
    return { ok: true, queued: true };
  }

  async function post(entry) {
    if (entry.kind === "photo") {
      const blob = await (await fetch(entry.dataUrl)).blob();
      const fd = new FormData();
      fd.append("file", blob, "photo.jpg");
      const p = new URLSearchParams({ client_event_id: entry.id, at: entry.at, caption: entry.caption || "", ...(entry.query || {}) });
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
      toast(N(sent, "Sent {n} saved update", "Sent {n} saved updates"), "ok");
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
    const waiting = q.length - refused.length;
    if (offline) parts.push(esc(T("No signal — you can keep working.")));
    if (waiting) parts.push(esc(N(waiting, "{n} update saved on this phone.", "{n} updates saved on this phone.")) + (syncing ? ` ${esc(T("(sending…)"))}` : ""));
    if (refused.length) parts.push(`${esc(N(refused.length, "{n} could not be saved:", "{n} could not be saved:"))} ${esc(refused[0].error)}`);
    bar.innerHTML = `<span>${parts.join(" ")}</span>${q.length && !offline ? `<button type="button" class="tp-link" data-act="sync">${esc(T("Send now"))}</button>` : ""}${refused.length ? `<button type="button" class="tp-link" data-act="dropRefused">${esc(T("Clear"))}</button>` : ""}`;
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
      throw new Error(T("No signal and nothing saved on this phone yet."));
    }
  }

  function staleNote(r) {
    return r.fresh ? "" : `<div class="tp-stale">${esc(T("Showing what was saved at {t} — no signal.", { t: dayTime(r.cachedAt) }))}</div>`;
  }

  function pendingFor(woId) {
    return queue().filter((x) => Number(x.woId) === Number(woId) && !x.error);
  }

  // ------------------------------------------------------------ navigation
  function go(name, params = {}) {
    view = { name, ...params };
    const url = new URL(window.location.href);
    ["wo", "asset", "tab", "inspect"].forEach((k) => url.searchParams.delete(k));
    if (name === "job") url.searchParams.set("wo", params.id);
    if (name === "asset") url.searchParams.set("asset", params.code);
    if (name === "inspectForm") url.searchParams.set("inspect", params.code);
    if (["shift", "week", "inspect"].includes(name)) url.searchParams.set("tab", name);
    history.replaceState(null, "", url);
    render();
    window.scrollTo(0, 0);
  }

  function setTabs() {
    const tab = { job: "today", inspectForm: "inspect", asset: "scan", available: "today" }[view.name] || view.name;
    document.querySelectorAll(".tp-tab").forEach((b) => b.classList.toggle("on", b.dataset.go === tab));
  }

  async function render() {
    setTabs();
    clearInterval(tick);
    const main = qs("tpMain");
    if (!main) return;
    if (!main.innerHTML.trim()) main.innerHTML = `<div class="tp-empty">${esc(T("Loading…"))}</div>`;
    try {
      if (view.name === "today") await renderToday(main);
      else if (view.name === "job") await renderJob(main);
      else if (view.name === "shift") await renderShift(main);
      else if (view.name === "week") await renderWeek(main);
      else if (view.name === "asset") await renderAsset(main);
      else if (view.name === "scan") renderScan(main);
      else if (view.name === "available") await renderAvailable(main);
      else if (view.name === "inspect") await renderInspect(main);
      else if (view.name === "inspectForm") await renderInspectForm(main);
    } catch (e) {
      main.innerHTML = `<div class="tp-card tp-error">${esc(e.message || e)}</div><button class="tp-btn" data-go="today">${esc(T("Back to Today"))}</button>`;
    }
    renderSyncBar();
  }

  // ------------------------------------------------------------ Today
  function jobCard(c) {
    const st = STATE_CLASS[c.my_state] || "idle";
    const pend = pendingFor(c.id).length;
    const partsLine = c.parts?.short ? N(c.parts.short, "{n} part short", "{n} parts short") : c.parts?.open_requests ? N(c.parts.open_requests, "{n} part request open", "{n} part requests open") : "";
    return `
      <button type="button" class="tp-job ${c.critical ? "crit" : ""}" data-open="${c.id}">
        <div class="tp-job-top">
          <span class="tp-asset">${esc(c.asset_code)}</span>
          <span class="tp-chip ${st}">${esc(stateText(c.my_state))}</span>
        </div>
        <div class="tp-job-line">${esc(c.job)}</div>
        <div class="tp-job-meta">
          <span>${esc(T("WO #{id}", { id: c.id }))}</span>
          ${c.role === "helper" ? `<span>${esc(T("Helper"))}</span>` : ""}
          ${c.due_date ? `<span>${esc(T("Due {d}", { d: String(c.due_date).slice(0, 10) }))}</span>` : ""}
          ${c.my_hours_today ? `<span>${esc(T("{h} today", { h: hrs(c.my_hours_today) }))}</span>` : ""}
          ${partsLine ? `<span class="warn">${esc(partsLine)}</span>` : ""}
          ${pend ? `<span class="warn">${esc(T("{n} not sent", { n: pend }))}</span>` : ""}
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
          <div class="tp-muted">${esc(d.shift ? T("Shift since {t} · {h} on jobs today", { t: time(d.shift.started_at), h: hrs(d.hours_today) }) : T("Shift not started"))}</div>
        </div>
        ${d.shift ? `<button type="button" class="tp-btn sm" data-go="shift">${esc(T("Shift report"))}</button>` : `<button type="button" class="tp-btn sm primary" data-act="shiftStart">${esc(T("Start shift"))}</button>`}
      </div>
      ${d.shift?.overdue ? `
        <div class="tp-note warn-shift" data-go="shift">
          <b>${esc(T("Your shift report from {t} is not submitted yet.", { t: dayTime(d.shift.started_at) }))}</b>
          <div class="small">${esc(T("Submit it now so your job time reaches the timesheet. Your next job then starts a new shift."))}</div>
        </div>` : ""}
      ${d.other_running ? `
        <div class="tp-current other">
          <div class="tp-muted small">${esc(T("Other time (not on a machine)"))}</div>
          <div class="tp-current-asset">${esc(reasonText(d.other_running.reason))}</div>
          ${d.other_running.note ? `<div>${esc(d.other_running.note)}</div>` : ""}
          <div class="tp-timer" data-since="${esc(d.other_running.since)}">${esc(since(d.other_running.since))}</div>
          <div class="tp-row"><button type="button" class="tp-btn" data-act="otherStop">${esc(T("Stop"))}</button></div>
        </div>` : ""}
      ${cur ? `
        <div class="tp-current" data-open="${cur.id}">
          <div class="tp-muted small">${esc(T(cur.state === "testing" ? "Testing now" : "Working on now"))}</div>
          <div class="tp-current-asset">${esc(cur.asset_code)} <span class="tp-muted">${esc(T("WO #{id}", { id: cur.id }))}</span></div>
          <div>${esc(cur.job)}</div>
          <div class="tp-timer" data-since="${esc(cur.since)}">${esc(since(cur.since))}</div>
          <div class="tp-row">
            <button type="button" class="tp-btn" data-quick="pause" data-wo="${cur.id}">${esc(T("Pause"))}</button>
            <button type="button" class="tp-btn primary" data-open="${cur.id}">${esc(T("Open job"))}</button>
          </div>
        </div>` : ""}
      ${(d.notifications || []).map((n) => `<div class="tp-note ${esc(n.kind)}" data-open="${n.wo_id}">${esc(noteText(n))}</div>`).join("")}
      <div class="tp-stats">
        <div><b>${d.counts.urgent}</b><span>${esc(T("Urgent"))}</span></div>
        <div><b>${d.counts.planned}</b><span>${esc(T("Planned"))}</span></div>
        <div><b>${d.counts.waiting}</b><span>${esc(T("Waiting"))}</span></div>
        <div><b>${d.counts.completed}</b><span>${esc(T("Done today"))}</span></div>
      </div>
      <div class="tp-row">
        <button type="button" class="tp-btn grow" data-act="otherForm">${esc(T("Other time"))}</button>
        <button type="button" class="tp-btn grow" data-go="scan">${esc(T("Work on a machine"))}</button>
      </div>
      ${d.available_count ? `<button type="button" class="tp-btn full tp-pickup" data-go="available">${esc(N(d.available_count, "Pick up a job — {n} open job nobody has taken", "Pick up a job — {n} open jobs nobody has taken"))}</button>` : ""}
      ${section(T("Urgent"), g.urgent)}
      ${section(T("Planned"), g.planned)}
      ${section(T("Waiting"), g.waiting)}
      ${nothing ? `<div class="tp-card tp-empty">${esc(T("No open jobs for you. Your foreman assigns jobs on the Work Orders board."))}</div>` : ""}
      ${section(T("Completed today"), g.completed)}
      <div class="tp-card">
        <div class="tp-row">
          <input id="tpWoNo" type="number" inputmode="numeric" min="1" placeholder="${esc(T("Open WO #"))}" class="tp-input grow" />
          <button type="button" class="tp-btn" data-act="openWoNo">${esc(T("Open"))}</button>
        </div>
      </div>`;
    startTimers();
  }

  function noteText(n) {
    if (n.kind === "assigned" && n.asset_code) return T("New job: WO #{id} {asset} — {job}", { id: n.wo_id, asset: n.asset_code, job: n.job || "" });
    if (n.kind === "parts" && n.part) return T("Part arrived for WO #{id}: {part}", { id: n.wo_id, part: n.part });
    return n.text;
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
        <button type="button" class="tp-back" data-go="today">‹ ${esc(T("Today"))}</button>
        <div class="tp-jobhead-main">
          <div class="tp-asset lg">${esc(d.asset.code)} <span class="tp-muted">${esc(d.asset.name || "")}</span></div>
          <div class="tp-job-line">${esc(d.wo.job)}</div>
          <div class="tp-job-meta">
            <span>${esc(T("WO #{id}", { id: d.wo.id }))}</span>
            <span>${esc(woStatus(d.wo.status))}</span>
            ${Number(d.asset.meter?.hours) > 0 ? `<span>${esc(T("{h} h on meter", { h: Number(d.asset.meter.hours).toFixed(0) }))}</span>` : ""}
            ${me.role !== "lead" ? `<span>${esc(T(me.role === "helper" ? "You are helping" : "Viewing as foreman"))}</span>` : ""}
          </div>
        </div>
      </div>
      <div class="tp-state ${STATE_CLASS[state] || "idle"}">
        <div>
          <div class="tp-state-label">${esc(finished ? T("Job completed") : stateText(state))}</div>
          <div class="tp-muted small">
            ${esc(T("You: {a} · Everyone: {b}", { a: hrs(d.time.my_hours), b: hrs(d.time.total_hours) }))}
            ${running && d.time.running_since ? ` · ${esc(T("running"))} <span data-since="${esc(d.time.running_since)}">${esc(since(d.time.running_since))}</span>` : ""}
          </div>
        </div>
        ${pend.length ? `<span class="tp-chip waiting">${esc(T("{n} not sent", { n: pend.length }))}</span>` : ""}
      </div>
      ${actions.length ? `<div class="tp-actions">${actions.map(([a, label, cls]) => `<button type="button" class="tp-btn ${cls}" data-action="${a}">${esc(T(label))}</button>`).join("")}</div>` : ""}
      ${!finished && String(d.wo.status).toLowerCase() === "open" ? `<div class="tp-card tp-muted">${esc(T("This job is not assigned yet. Ask your foreman to assign it."))}</div>` : ""}
      <div class="tp-subtabs">
        ${[["work", T("Job")], ["parts", `${T("Parts")}${d.parts.summary.short || d.parts.summary.open_requests ? " •" : ""}`], ["notes", `${T("Notes & photos")}${d.findings.length + d.photos.length ? ` (${d.findings.length + d.photos.length})` : ""}`], ["history", T("History")], ["borris", T("Ask Borris")]]
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
        <div class="tp-h4">${esc(T("Breakdown"))} ${b.critical ? `<span class="tp-chip bad">${esc(T("Critical"))}</span>` : ""}</div>
        <div>${esc([b.component, b.description].filter(Boolean).join(" — "))}</div>
        <div class="tp-muted small">${b.since ? esc(T("Down since {t}", { t: dayTime(b.since) })) : ""}${b.return_date ? ` · ${esc(T("Return date {d}", { d: b.return_date }))}` : ""}${b.parts_status ? ` · ${esc(T("Parts: {s}", { s: b.parts_status }))}` : ""}</div>
      </div>` : ""}
      ${s ? `<div class="tp-card">
        <div class="tp-h4">${esc(T("Service"))}</div>
        <div>${esc(s.name)} · ${esc(T("every {n} h", { n: s.interval }))}</div>
        <div class="tp-muted small">${esc(T("Due at {n} h", { n: s.next_due }))} · ${esc(dueText(s.remaining))}</div>
      </div>` : ""}
      ${d.wo.job_description ? `<div class="tp-card"><div class="tp-h4">${esc(T("Job card"))}</div><div class="tp-pre">${esc(d.wo.job_description)}</div></div>` : ""}
      ${d.wo.progress ? `<div class="tp-card"><div class="tp-h4">${esc(T("Latest progress"))}</div><div>${esc(d.wo.progress)}</div></div>` : ""}
      ${d.wo.completion_notes ? `<div class="tp-card"><div class="tp-h4">${esc(T("What was done"))}</div><div>${esc(d.wo.completion_notes)}</div></div>` : ""}
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Team"))}</div>
        <div>${esc(T("Lead: {name}", { name: team.lead_name || team.lead || T("not assigned") }))}</div>
        ${(team.helpers || []).map((h) => `<div>${esc(T("Helper: {name}", { name: h.name }))} <span class="tp-muted small">${esc(stateText(team.states?.[h.username] || "idle"))}</span>
          ${isLead() ? `<button type="button" class="tp-link" data-act="helperDel" data-user="${esc(h.username)}">${esc(T("Remove"))}</button>` : ""}</div>`).join("")}
        ${isLead() ? `<div class="tp-row mt"><input id="tpHelper" class="tp-input grow" placeholder="${esc(T("Add helper (username)"))}" /><button type="button" class="tp-btn sm" data-act="helperAdd">${esc(T("Add"))}</button></div>` : ""}
      </div>`;
  }

  function dueText(remaining) {
    return remaining >= 0 ? T("{n} h to go", { n: remaining }) : T("{n} h overdue", { n: Math.abs(remaining) });
  }

  function partsTab(d) {
    const p = d.parts;
    const stateChip = { issued: ["ok", T("Issued")], in_stock: ["active", T("In stores")], short: ["bad", T("Short")] };
    const reqStatus = { requested: "Requested", ordered: "Ordered", received: "Received", cancelled: "Cancelled", rejected: "Rejected" };
    return `
      ${p.planned.length ? `<div class="tp-card"><div class="tp-h4">${esc(T("Planned for this job"))}</div>
        ${p.planned.map((l) => `<div class="tp-part">
          <div><b>${esc(l.part_code || "")}</b> ${esc(l.description || "")}</div>
          <div class="tp-muted small">${esc(T("Need {q} {u} · issued {i} · in stores {s}", { q: l.quantity_planned, u: l.unit_of_measure || "", i: l.quantity_issued, s: l.on_hand == null ? "?" : l.on_hand }))}${l.bin ? ` · ${esc(T("bin {b}", { b: l.bin }))}` : ""}</div>
          <span class="tp-chip ${stateChip[l.state][0]}">${esc(stateChip[l.state][1])}</span>
        </div>`).join("")}</div>` : ""}
      ${p.issued.length ? `<div class="tp-card"><div class="tp-h4">${esc(T("Issued to this job"))}</div>
        ${p.issued.map((l) => `<div class="tp-part"><div><b>${esc(l.part_code)}</b> ${esc(l.part_name)}</div><div class="tp-muted small">${esc(T("{n} issued", { n: l.qty }))}${l.bin ? ` · ${esc(T("from {b}", { b: l.bin }))}` : ""}</div></div>`).join("")}</div>` : ""}
      <div class="tp-card"><div class="tp-h4">${esc(T("Requests to stores"))}</div>
        ${p.requests.length ? p.requests.map((r) => `<div class="tp-part"><div><b>${esc(r.part_code || "")}</b> ${esc(r.part_name || "")} × ${esc(r.qty)}</div>
          <div class="tp-muted small">${esc(r.requested_by || "")} · ${esc(dayTime(r.created_at))}${r.status_notes ? ` · ${esc(r.status_notes)}` : ""}</div>
          <span class="tp-chip ${String(r.status).toLowerCase() === "received" ? "ok" : String(r.status).toLowerCase() === "ordered" ? "active" : "waiting"}">${esc(T(reqStatus[String(r.status || "requested").toLowerCase()] || r.status))}</span></div>`).join("") : `<div class="tp-muted">${esc(T("No requests yet."))}</div>`}
      </div>
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Request a part"))}</div>
        <p class="tp-muted small">${esc(T("Stores issue parts to the job. Your request goes to the stores queue."))}</p>
        <input id="tpPartQ" class="tp-input" placeholder="${esc(T("Search stores (code or name)"))}" autocomplete="off" />
        <div id="tpPartHits"></div>
        <input id="tpPartCode" class="tp-input" placeholder="${esc(T("Part code (if known)"))}" />
        <input id="tpPartName" class="tp-input" placeholder="${esc(T("What part is needed"))}" />
        <div class="tp-row">
          <input id="tpPartQty" class="tp-input" type="number" inputmode="decimal" min="1" value="1" style="max-width:90px" />
          <select id="tpPartUrg" class="tp-input"><option value="normal">${esc(T("Normal"))}</option><option value="urgent">${esc(T("Urgent — machine down"))}</option></select>
        </div>
        <input id="tpPartNote" class="tp-input" placeholder="${esc(T("Note for stores (optional)"))}" />
        <button type="button" class="tp-btn primary full" data-act="partRequest">${esc(T("Send to stores"))}</button>
      </div>`;
  }

  function notesTab(d) {
    return `
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Add a finding"))}</div>
        <textarea id="tpFinding" class="tp-input" rows="3" placeholder="${esc(T("What did you find? e.g. hose chafed on the chassis clamp"))}"></textarea>
        <div class="tp-row">
          <select id="tpFindingKind" class="tp-input"><option value="finding">${esc(T("Finding"))}</option><option value="safety">${esc(T("Safety issue"))}</option><option value="note">${esc(T("Note"))}</option></select>
          <button type="button" class="tp-btn primary" data-act="finding">${esc(T("Save"))}</button>
        </div>
        <label class="tp-btn full tp-photo-btn">📷 ${esc(T("Add photo"))}<input id="tpPhoto" type="file" accept="image/*" capture="environment" hidden /></label>
      </div>
      ${d.photos.length ? `<div class="tp-photos">${d.photos.map((p) => `<a href="${esc(p.url)}" target="_blank" rel="noopener"><img src="${esc(p.url)}" alt="${esc(p.caption || T("Job photo"))}" loading="lazy" /></a>`).join("")}</div>` : ""}
      ${d.findings.map((f) => `<div class="tp-card tp-finding ${esc(f.kind)}"><div class="tp-muted small">${esc(T(f.kind === "safety" ? "Safety" : f.kind === "note" ? "Note" : "Finding"))} · ${esc(f.username)} · ${esc(dayTime(f.at))}</div><div>${esc(f.text)}</div></div>`).join("")}`;
  }

  function historyTab(d) {
    return `
      <div class="tp-card"><div class="tp-h4">${esc(T("Earlier jobs on {code}", { code: d.asset.code }))}</div>
        ${d.history.length ? d.history.map((h) => `<div class="tp-hist"><div><b>${esc(T("WO #{id}", { id: h.id }))}</b> ${esc(h.job)}</div><div class="tp-muted small">${esc(dayTime(h.done_at))}${h.notes ? ` · ${esc(h.notes)}` : ""}</div></div>`).join("") : `<div class="tp-muted">${esc(T("No earlier jobs recorded."))}</div>`}
      </div>
      <div class="tp-card"><div class="tp-h4">${esc(T("Activity on this job"))}</div>
        ${d.activity.length ? d.activity.map((a) => `<div class="tp-hist"><div>${esc(a.text)}</div><div class="tp-muted small">${esc(a.who)} · ${esc(dayTime(a.at))}${a.auto ? ` · ${esc(T("automatic"))}` : ""}</div></div>`).join("") : `<div class="tp-muted">${esc(T("Nothing yet."))}</div>`}
      </div>
      <button type="button" class="tp-btn full" data-asset="${esc(d.asset.code)}">${esc(T("Machine view: {code}", { code: d.asset.code }))}</button>`;
  }

  function borrisTab(d) {
    const ideas = (d.breakdown
      ? ["What usually causes this failure?", "What should I check first?", "Which manual section covers this?"]
      : ["What does this service include?", "Which manual section covers this?", "What torque or fluid spec applies?"]).map((q) => T(q));
    return `
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Ask Borris about this job"))}</div>
        <p class="tp-muted small">${esc(T("Borris gives advice only. He cannot change the job, stock or the service plan."))}</p>
        <div class="tp-ideas">${ideas.map((q) => `<button type="button" class="tp-idea" data-idea="${esc(q)}">${esc(q)}</button>`).join("")}</div>
        <textarea id="tpAsk" class="tp-input" rows="2" placeholder="${esc(T("Your question"))}"></textarea>
        <button type="button" class="tp-btn primary full" data-act="ask">${esc(T("Ask"))}</button>
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

  const ACTION_STATE = { start: "active", resume: "active", pause: "paused", waiting_parts: "waiting_parts", waiting_ops: "waiting_ops", testing: "testing", complete: "done" };

  async function doAction(woId, action, extra = {}) {
    const res = await send({ kind: "action", woId, url: `/workorders/${woId}/action`, body: { action, ...extra }, state: ACTION_STATE[action] });
    const label = stateText(ACTION_STATE[action]);
    toast(res.queued ? T("{s} — saved on this phone, will send when there is signal", { s: label }) : label, res.queued ? "" : "ok");
    return res;
  }

  function askAction(woId, action, d) {
    if (action === "complete") {
      const lead = d?.me?.role !== "helper";
      sheet(`
        <h3>${esc(T(lead ? "Complete the job" : "Finish your part"))}</h3>
        ${lead ? `<label class="tp-label">${esc(T("What was done"))} <b>${esc(T("(required)"))}</b></label>
        <textarea id="tpDone" class="tp-input" rows="4" placeholder="${esc(T("e.g. Replaced hydraulic hose and clamp, tested under load, no leaks"))}"></textarea>
        <label class="tp-label">${esc(T("Labour hours (leave empty to use the job timer: {h})", { h: hrs(d?.time?.total_hours || 0) }))}</label>
        <input id="tpDoneHours" class="tp-input" type="number" inputmode="decimal" min="0" step="0.25" placeholder="${esc(Number(d?.time?.total_hours || 0).toFixed(2))}" />` : `<p class="tp-muted">${esc(T("The lead technician completes the work order. This stops your time on it."))}</p>
        <textarea id="tpDone" class="tp-input" rows="2" placeholder="${esc(T("What you did (optional)"))}"></textarea>`}
        <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">${esc(T("Cancel"))}</button><button type="button" class="tp-btn ok grow" data-act="confirmComplete" data-wo="${woId}" data-lead="${lead ? 1 : 0}">${esc(T(lead ? "Complete job" : "Finish my part"))}</button></div>`);
      return;
    }
    if (action === "waiting_parts" || action === "waiting_ops" || action === "pause") {
      const title = T({ waiting_parts: "Waiting for parts", waiting_ops: "Waiting for operations", pause: "Pause the job" }[action]);
      const hint = T({ waiting_parts: "Which part? e.g. hose from stores, ordered from Maputo", waiting_ops: "e.g. machine needed on the bench, waiting for the operator", pause: "Why? (optional) e.g. lunch, called to another job" }[action]);
      sheet(`
        <h3>${esc(title)}</h3>
        <textarea id="tpWhy" class="tp-input" rows="2" placeholder="${esc(hint)}"></textarea>
        <p class="tp-muted small">${esc(T("Time stops counting as labour until you resume."))}</p>
        <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">${esc(T("Cancel"))}</button><button type="button" class="tp-btn primary grow" data-act="confirmWhy" data-wo="${woId}" data-action="${action}">${esc(title)}</button></div>`);
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
        <h2 class="tp-title">${esc(T("Shift report"))}</h2>
        <div class="tp-card"><p>${esc(T("Your shift has not started. It starts when you tap Start shift or start your first job, and ends when you submit the report — a night shift over midnight is one report."))}</p>
        <button type="button" class="tp-btn primary big full" data-act="shiftStart">${esc(T("Start shift"))}</button></div>
        ${isLead() ? `<button type="button" class="tp-btn full" data-act="teamShifts">${esc(T("Team shift reports"))}</button><div id="tpTeam"></div>` : ""}`;
      return;
    }
    const labour = s.timeline.filter((t) => t.labour);
    main.innerHTML = `${staleNote(r)}
      <h2 class="tp-title">${esc(T("Shift report"))}</h2>
      <div class="tp-card">
        <div class="tp-row between"><div>${esc(T("Started {t}", { t: dayTime(s.started_at) }))}</div><div class="tp-big">${esc(hrs(s.hours))}</div></div>
        <div class="tp-muted small">${esc(T("on jobs"))} · ${esc(N(labour.length, "{n} work block", "{n} work blocks"))} · ${esc(N(s.findings_logged.length, "{n} finding logged", "{n} findings logged"))}</div>
      </div>
      <div class="tp-card"><div class="tp-h4">${esc(T("What you did"))}</div>
        ${s.timeline.length ? s.timeline.map((t) => `<div class="tp-tl ${STATE_CLASS[t.state] || ""}">
          <div class="tp-tl-time">${esc(time(t.start))}–${t.open ? esc(T("now")) : esc(time(t.end))}</div>
          <div><b>${esc(t.asset_code || "")}</b> ${esc(t.job || "")} <span class="tp-muted small">${esc(T("WO #{id}", { id: t.work_order_id }))}</span><div class="tp-muted small">${esc(stateText(t.state))} · ${esc(hrs(t.hours))}</div></div>
        </div>`).join("") : `<div class="tp-muted">${esc(T("No job time yet this shift."))}</div>`}
      </div>
      <div class="tp-card"><div class="tp-row between"><div class="tp-h4">${esc(T("Other time"))}</div><div>${esc(hrs(s.other_hours || 0))}</div></div>
        ${(s.other || []).length ? s.other.map((o) => `<div class="tp-tl paused"><div class="tp-tl-time">${esc(time(o.start))}–${o.open ? esc(T("now")) : esc(time(o.end))}</div><div><b>${esc(reasonText(o.reason))}</b>${o.note ? ` <span class="tp-muted small">${esc(o.note)}</span>` : ""}<div class="tp-muted small">${esc(hrs(o.hours))}</div></div></div>`).join("") : `<div class="tp-muted small">${esc(T("Housekeeping, training, meetings, travel or waiting — time that is not on a machine."))}</div>`}
        <button type="button" class="tp-btn full mt" data-act="otherForm">${esc(T("Log other time"))}</button>
      </div>
      ${s.findings_logged.length ? `<div class="tp-card"><div class="tp-h4">${esc(T("Findings logged on jobs"))}</div>${s.findings_logged.map((f) => `<div class="tp-hist">${esc(f.text)} <span class="tp-muted small">${f.work_order_id ? `${esc(T("WO #{id}", { id: f.work_order_id }))}` : ""} ${esc(time(f.at))}</span></div>`).join("")}</div>` : ""}
      <div class="tp-card">
        ${SHIFT_FIELDS.map(([k, label, hint]) => `<label class="tp-label">${esc(T(label))}</label><textarea class="tp-input" rows="2" data-shift="${k}" placeholder="${esc(T(hint))}">${esc(s[k] || "")}</textarea>`).join("")}
        <div class="tp-muted small">${esc(T("Saved as you type."))}</div>
      </div>
      <button type="button" class="tp-btn ok big full" data-act="shiftSubmit">${esc(T("Submit shift report"))}</button>
      ${isLead() ? `<button type="button" class="tp-btn full" data-act="teamShifts">${esc(T("Team shift reports"))}</button><div id="tpTeam"></div>` : ""}`;
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
    host.innerHTML = `<div class="tp-muted">${esc(T("Loading…"))}</div>`;
    const d = await A.fetchJson(api("/shifts?days=7"));
    host.innerHTML = d.rows.length ? d.rows.map((s) => `<details class="tp-card">
      <summary><b>${esc(s.name)}</b> · ${esc(dayTime(s.started_at))} → ${esc(time(s.ended_at))} · ${esc(hrs(s.hours))}</summary>
      ${s.timeline.filter((t) => t.labour).map((t) => `<div class="tp-muted small">${esc(time(t.start))}–${esc(time(t.end))} ${esc(t.asset_code || "")} ${esc(T("WO #{id}", { id: t.work_order_id }))} ${esc(t.job || "")} (${esc(hrs(t.hours))})</div>`).join("")}
      ${SHIFT_FIELDS.filter(([k]) => s[k]).map(([k, label]) => `<div class="mt"><b>${esc(T(label))}:</b> ${esc(s[k])}</div>`).join("")}
    </details>`).join("") : `<div class="tp-muted">${esc(T("No submitted reports in the last 7 days."))}</div>`;
  }

  // ------------------------------------------------------------ Week
  async function renderWeek(main) {
    const r = await load("week", "/week");
    const d = r.data;
    const max = Math.max(8, ...d.days.map((x) => x.hours));
    main.innerHTML = `${staleNote(r)}
      <h2 class="tp-title">${esc(T("My week"))}</h2>
      <div class="tp-stats">
        <div><b>${esc(Number(d.hours).toFixed(1))}</b><span>${esc(T("Hours on jobs"))}</span></div>
        <div><b>${d.jobs_completed}</b><span>${esc(T("Completed"))}</span></div>
        <div><b>${d.jobs_open}</b><span>${esc(T("Open"))}</span></div>
        <div><b>${d.jobs_waiting}</b><span>${esc(T("Waiting"))}</span></div>
      </div>
      <div class="tp-card">
        ${d.days.map((x) => `<div class="tp-bar"><span class="tp-bar-day">${esc(new Date(`${x.day}T12:00:00`).toLocaleDateString(locale(), { weekday: "short", day: "2-digit" }))}</span><span class="tp-bar-track"><span style="width:${Math.round((x.hours / max) * 100)}%"></span></span><span class="tp-bar-val">${esc(hrs(x.hours))}</span></div>`).join("")}
      </div>
      <div class="tp-card tp-muted small">${esc(N(d.shifts_submitted, "{n} shift report submitted this week", "{n} shift reports submitted this week"))}${d.shift_open ? ` · ${esc(T("a shift is open now"))}` : ""}.</div>`;
  }

  // ------------------------------------------------------------ Pick up a job
  async function renderAvailable(main) {
    const r = await load("available", "/available");
    const rows = r.data.rows || [];
    main.innerHTML = `${staleNote(r)}
      <div class="tp-jobhead"><button type="button" class="tp-back" data-go="today">‹ ${esc(T("Today"))}</button>
        <div class="tp-jobhead-main"><h2 class="tp-title">${esc(T("Pick up a job"))}</h2>
        <div class="tp-muted small">${esc(T("Open jobs nobody has taken yet. When you take one it is assigned to you and your foreman sees it on the board."))}</div></div></div>
      ${rows.length ? rows.map((c) => `
        <div class="tp-job ${c.critical ? "crit" : ""}">
          <div class="tp-job-top">
            <span class="tp-asset">${esc(c.asset_code)}</span>
            ${c.critical ? `<span class="tp-chip bad">${esc(T("Critical"))}</span>` : c.urgent ? `<span class="tp-chip waiting">${esc(T("Urgent"))}</span>` : ""}
          </div>
          <div class="tp-job-line">${esc(c.job)}</div>
          <div class="tp-job-meta">
            <span>${esc(T("WO #{id}", { id: c.id }))}</span>
            ${c.opened_at ? `<span>${esc(T("Opened {t}", { t: dayTime(c.opened_at) }))}</span>` : ""}
            ${c.parts?.short ? `<span class="warn">${esc(N(c.parts.short, "{n} part short", "{n} parts short"))}</span>` : ""}
          </div>
          <div class="tp-row mt"><button type="button" class="tp-btn primary grow" data-act="claim" data-wo="${c.id}" data-label="${esc(`${c.asset_code} — ${c.job}`)}">${esc(T("Take this job"))}</button></div>
        </div>`).join("") : `<div class="tp-card tp-empty">${esc(T("No open jobs waiting. Everything is assigned."))}</div>`}`;
  }

  // ------------------------------------------------------------ Machine (QR)
  function renderScan(main) {
    main.innerHTML = `
      <h2 class="tp-title">${esc(T("Machine"))}</h2>
      <div class="tp-card">
        <p class="tp-muted">${esc(T("Scan the machine's QR label with the phone camera, or type the fleet number."))}</p>
        <div class="tp-row"><input id="tpAssetCode" class="tp-input grow" placeholder="${esc(T("e.g. T01AM"))}" autocapitalize="characters" /><button type="button" class="tp-btn primary" data-act="openAsset">${esc(T("Open"))}</button></div>
      </div>`;
  }

  async function renderAsset(main) {
    const code = String(view.code || "").toUpperCase();
    const r = await load(`asset:${code}`, `/assets/${encodeURIComponent(code)}`);
    const d = r.data;
    const a = d.asset;
    main.innerHTML = `${staleNote(r)}
      <div class="tp-jobhead"><button type="button" class="tp-back" data-go="scan">‹ ${esc(T("Machines"))}</button>
        <div class="tp-jobhead-main"><div class="tp-asset lg">${esc(a.asset_code)} <span class="tp-muted">${esc(a.asset_name || "")}</span></div>
        <div class="tp-job-meta">${Number(a.meter?.hours) > 0 ? `<span>${esc(T("{h} h on meter", { h: Number(a.meter.hours).toFixed(0) }))}</span>` : ""}${d.breakdown ? `<span class="bad">${esc(T("Broken down"))}</span>` : `<span>${esc(T("Running"))}</span>`}</div></div></div>
      ${d.breakdown ? `<div class="tp-card"><div class="tp-h4">${esc(T("Open breakdown"))} ${d.breakdown.critical ? `<span class="tp-chip bad">${esc(T("Critical"))}</span>` : ""}</div><div>${esc([d.breakdown.component, d.breakdown.description].filter(Boolean).join(" — "))}</div><div class="tp-muted small">${esc(T("Down since {t}", { t: dayTime(d.breakdown.start_at || d.breakdown.breakdown_date) }))}${d.breakdown.ets_repair_date ? ` · ${esc(T("Return date {d}", { d: d.breakdown.ets_repair_date }))}` : ""}</div></div>` : ""}
      <div class="tp-card"><div class="tp-h4">${esc(T("Open jobs"))}</div>
        ${d.active.length ? d.active.map((w) => `<button type="button" class="tp-job" data-open="${w.id}"><div class="tp-job-top"><span>${esc(T("WO #{id}", { id: w.id }))}</span><span class="tp-chip ${w.mine ? "active" : "idle"}">${w.mine ? esc(T("Yours")) : esc(w.lead || T("Unassigned"))}</span></div><div class="tp-job-line">${esc(w.job)}</div></button>`).join("") : `<div class="tp-muted">${esc(T("No open jobs on this machine."))}</div>`}
      </div>
      ${d.next_service ? `<div class="tp-card"><div class="tp-h4">${esc(T("Next service"))}</div><div>${esc(T("{s} at {n} h", { s: d.next_service.service, n: d.next_service.due_at }))}</div><div class="tp-muted small">${esc(dueText(d.next_service.remaining))}</div></div>` : ""}
      <div class="tp-card"><div class="tp-h4">${esc(T("Recent work"))}</div>${d.history.length ? d.history.map((h) => `<div class="tp-hist"><div><b>${esc(T("WO #{id}", { id: h.id }))}</b> ${esc(h.job)}</div><div class="tp-muted small">${esc(dayTime(h.done_at))}${h.notes ? ` · ${esc(h.notes)}` : ""}</div></div>`).join("") : `<div class="tp-muted">${esc(T("No completed jobs yet."))}</div>`}</div>
      ${d.findings.length ? `<div class="tp-card"><div class="tp-h4">${esc(T("Findings"))}</div>${d.findings.map((f) => `<div class="tp-hist">${esc(f.text)} <span class="tp-muted small">${esc(f.username)} · ${esc(dayTime(f.at))}</span></div>`).join("")}</div>` : ""}
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Found a problem?"))}</div>
        <textarea id="tpAssetFinding" class="tp-input" rows="2" placeholder="${esc(T("What did you find?"))}"></textarea>
        <div class="tp-row"><button type="button" class="tp-btn grow" data-act="assetFinding" data-asset-id="${a.id}">${esc(T("Save finding"))}</button><button type="button" class="tp-btn warn grow" data-act="breakdownForm" data-code="${esc(a.asset_code)}">${esc(T("Report breakdown"))}</button></div>
      </div>
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Inspection"))}</div>
        <p class="tp-muted small">${esc(T("Check the machine section by section, with photos. Faults open a work order and the next operator is told to bring it to the workshop."))}</p>
        <button type="button" class="tp-btn primary full" data-act="inspectStart" data-code="${esc(a.asset_code)}">${esc(T("Inspect this machine"))}</button>
      </div>
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Working on this machine without a work order?"))}</div>
        <p class="tp-muted small">${esc(T("For adjustments, greasing, quick fixes or helping an operator. A work order is opened for you and your time starts. It is not a breakdown."))}</p>
        <button type="button" class="tp-btn primary full" data-act="unplannedForm" data-code="${esc(a.asset_code)}">${esc(T("Log work on this machine"))}</button>
      </div>`;
  }


  // ------------------------------------------------------------ Inspection
  // Artisan inspection: every check OK / Fault / N/A (a fault needs a comment
  // and can have photos), details and the overall result. The form is kept on
  // the phone while it is filled in, and is sent later when there is no signal.
  const INSP_DRAFT = (code) => `ironlog-tech-insp:${code}`;
  let insp = null; // { code, answers, comments, details, photos: [{ id, key, dataUrl }], cid }
  let inspTpl = null;

  const pick = (o) => (lang() === "pt" ? o.pt || o.en : o.en);
  const other = (o) => (lang() === "pt" ? o.en : o.pt || "");
  const RESULT_CLASS = { fit: "ok", restricted: "waiting", not_fit: "bad" };

  function inspDraft(code) {
    if (insp && insp.code === code) return insp;
    const hour = new Date().getHours();
    insp = {
      code,
      answers: {},
      comments: {},
      photos: [],
      details: { type: "daily", shift: hour >= 6 && hour < 18 ? "day" : "night", hours: "", location: "", result: "", notes: "" },
      ...read(INSP_DRAFT(code), {}),
    };
    return insp;
  }

  function saveDraft() {
    if (!insp) return;
    if (!write(INSP_DRAFT(insp.code), insp)) toast(T("Phone storage is full: take fewer photos or send the inspection now."), "bad");
  }

  function dropDraft(code) {
    try { localStorage.removeItem(INSP_DRAFT(code)); } catch { /* ignore */ }
    if (insp?.code === code) insp = null;
  }

  async function openInspectionPdf(id) {
    const w = window.open("", "_blank");
    try {
      const res = await fetch(api(`/inspections/${id}/pdf`), { headers: A.authHeaders() });
      if (!res.ok) throw new Error(T("Could not open the PDF."));
      const url = URL.createObjectURL(await res.blob());
      if (w) w.location.href = url;
      else window.location.href = url;
    } catch (e) {
      if (w) w.close();
      toast(e instanceof TypeError ? T("The PDF needs signal.") : e.message || String(e), "bad");
    }
  }

  function inspRow(r, withAsset = true) {
    const res = (inspTpl?.results || []).find((x) => x.key === r.overall_result);
    return `<div class="tp-hist">
      <div class="tp-job-top">${withAsset ? `<span class="tp-asset">${esc(r.asset_code)}</span>` : `<span>${esc(r.inspection_date)}</span>`}${res ? `<span class="tp-chip ${RESULT_CLASS[r.overall_result] || ""}">${esc(pick(res))}</span>` : ""}</div>
      <div class="tp-muted small">${withAsset ? `${esc(r.inspection_date)} · ` : ""}${esc(r.inspector_name || "")}${r.faults.length ? ` · ${esc(N(r.faults.length, "{n} fault", "{n} faults"))}` : ` · ${esc(T("No faults"))}`}${r.photo_count ? ` · ${esc(N(r.photo_count, "{n} photo", "{n} photos"))}` : ""}</div>
      <div class="tp-row mt">${r.work_order_id ? `<button type="button" class="tp-btn sm" data-open="${r.work_order_id}">${esc(T("WO #{id}", { id: r.work_order_id }))}</button>` : ""}<button type="button" class="tp-btn sm" data-act="inspPdf" data-id="${r.id}">${esc(T("PDF"))}</button></div>
    </div>`;
  }

  async function renderInspect(main) {
    const tpl = await load("insp:tpl", "/inspection-template").catch(() => null);
    if (tpl) inspTpl = tpl.data;
    let mine = null;
    try { mine = await load("insp:mine", "/inspections"); } catch { mine = null; }
    const drafts = Object.keys(localStorage).filter((k) => k.startsWith("ironlog-tech-insp:")).map((k) => k.split(":")[1]);
    main.innerHTML = `
      <h2 class="tp-title">${esc(T("Inspection"))}</h2>
      <div class="tp-card">
        <p class="tp-muted">${esc(T("Scan the machine's QR label, or type the fleet number."))}</p>
        <div class="tp-row"><input id="tpInspCode" class="tp-input grow" placeholder="${esc(T("e.g. T01AM"))}" autocapitalize="characters" /><button type="button" class="tp-btn primary" data-act="inspectOpen">${esc(T("Start"))}</button></div>
      </div>
      ${drafts.length ? `<div class="tp-card"><div class="tp-h4">${esc(T("Not sent yet"))}</div>${drafts.map((c) => `<button type="button" class="tp-btn full mt" data-act="inspectStart" data-code="${esc(c)}">${esc(T("Carry on with {code}", { code: c }))}</button>`).join("")}</div>` : ""}
      <div class="tp-card"><div class="tp-h4">${esc(T("My inspections"))}</div>
        ${mine ? staleNote(mine) : ""}
        ${mine?.data?.rows?.length ? mine.data.rows.map((r) => inspRow(r)).join("") : `<div class="tp-muted">${esc(T("No inspections yet."))}</div>`}
      </div>`;
  }

  function answerButtons(key, cur) {
    return `<div class="tp-ans">${[["ok", "OK"], ["fault", "Fault"], ["na", "N/A"]].map(([v, l]) => `<button type="button" class="tp-ans-b ${v} ${cur === v ? "on" : ""}" data-insp-ans="${esc(key)}" data-v="${v}" aria-pressed="${cur === v}">${esc(T(l))}</button>`).join("")}</div>`;
  }

  function photoStrip(key) {
    const list = insp.photos.filter((p) => (p.key || "") === key);
    return `<div class="tp-insp-photos">
      ${list.map((p) => `<div class="tp-insp-ph"><img src="${p.dataUrl}" alt="" /><button type="button" data-act="inspPhotoDel" data-id="${esc(p.id)}" aria-label="${esc(T("Remove photo"))}">×</button></div>`).join("")}
      <label class="tp-insp-add">📷 ${esc(T("Add photo"))}<input type="file" accept="image/*" capture="environment" data-insp-photo="${esc(key)}" hidden /></label>
    </div>`;
  }

  async function renderInspectForm(main) {
    const code = String(view.code || "").toUpperCase();
    const t = await load("insp:tpl", "/inspection-template");
    const tpl = t.data;
    inspTpl = tpl;
    let asset = null;
    try { asset = (await load(`asset:${code}`, `/assets/${encodeURIComponent(code)}`)).data.asset; } catch (e) {
      if (!isNetworkError(e) || e.status) throw e;
    }
    const d = inspDraft(code);
    const items = tpl.sections.flatMap((x) => x.items);
    const done = items.filter((i) => d.answers[i.key]).length;
    const faults = items.filter((i) => d.answers[i.key] === "fault").length;
    const meter = Number(asset?.meter?.hours) > 0 ? Number(asset.meter.hours).toFixed(0) : "";
    const y = window.scrollY;
    main.innerHTML = `
      <div class="tp-jobhead"><button type="button" class="tp-back" data-go="inspect">‹ ${esc(T("Inspection"))}</button>
        <div class="tp-jobhead-main"><div class="tp-asset lg">${esc(code)} <span class="tp-muted">${esc(asset?.asset_name || "")}</span></div>
        <div class="tp-job-meta"><span>${esc(T("{n} of {total} checked", { n: done, total: items.length }))}</span>${faults ? `<span class="bad">${esc(N(faults, "{n} fault", "{n} faults"))}</span>` : ""}</div></div></div>
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Details"))}</div>
        <label class="tp-label">${esc(T("Type of inspection"))}</label>
        <select class="tp-input" data-insp-d="type">${tpl.inspection_types.map((x) => `<option value="${x.key}" ${d.details.type === x.key ? "selected" : ""}>${esc(pick(x))}</option>`).join("")}</select>
        <label class="tp-label">${esc(T("Shift"))}</label>
        <div class="tp-ans">${["day", "night"].map((v) => `<button type="button" class="tp-ans-b ${d.details.shift === v ? "on" : ""}" data-insp-shift="${v}">${esc(T(v === "day" ? "Day" : "Night"))}</button>`).join("")}</div>
        <label class="tp-label">${esc(T("Hour meter"))}</label>
        <input class="tp-input" type="number" inputmode="decimal" min="0" step="0.1" data-insp-d="hours" value="${esc(d.details.hours)}" placeholder="${esc(meter ? T("Last known {h} h", { h: meter }) : "")}" />
        <label class="tp-label">${esc(T("Where is the machine?"))}</label>
        <input class="tp-input" data-insp-d="location" value="${esc(d.details.location)}" placeholder="${esc(T("e.g. Pit 2, plant, workshop bay 3"))}" />
      </div>
      ${tpl.sections.map((sec) => `
        <div class="tp-card">
          <div class="tp-h4">${esc(pick(sec))} <span class="tp-muted small">${esc(other(sec))}</span></div>
          ${sec.items.map((i) => {
            const a = d.answers[i.key] || "";
            return `<div class="tp-insp-item ${a === "fault" ? "fault" : ""}" id="tpi-${esc(i.key)}">
              <div class="tp-insp-label">${esc(pick(i))}<div class="tp-muted small">${esc(other(i))}</div></div>
              ${answerButtons(i.key, a)}
              ${a === "fault" ? `<textarea class="tp-input" rows="2" data-insp-note="${esc(i.key)}" placeholder="${esc(T("What is wrong? (required)"))}">${esc(d.comments[i.key] || "")}</textarea>${photoStrip(i.key)}` : ""}
            </div>`;
          }).join("")}
        </div>`).join("")}
      ${done < items.length ? `<button type="button" class="tp-btn full" data-act="inspAllOk">${esc(N(items.length - done, "Mark the {n} check left as OK", "Mark the {n} checks left as OK"))}</button>` : ""}
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Photos of the machine"))}</div>
        <p class="tp-muted small">${esc(T("Optional: overall view, meter, anything worth showing."))}</p>
        ${photoStrip("")}
      </div>
      <div class="tp-card">
        <div class="tp-h4">${esc(T("Overall result"))}</div>
        <div class="tp-results">${tpl.results.map((x) => `<button type="button" class="tp-result ${RESULT_CLASS[x.key]} ${d.details.result === x.key ? "on" : ""}" data-insp-result="${x.key}">${esc(pick(x))}</button>`).join("")}</div>
        <label class="tp-label">${esc(T("Notes"))}</label>
        <textarea class="tp-input" rows="3" data-insp-d="notes" placeholder="${esc(T("Anything else the workshop should know"))}">${esc(d.details.notes)}</textarea>
      </div>
      <div class="tp-row">
        <button type="button" class="tp-btn" data-act="inspDiscard" data-code="${esc(code)}">${esc(T("Discard"))}</button>
        <button type="button" class="tp-btn ok grow big" data-act="inspSubmit">${esc(T("Submit inspection"))}</button>
      </div>`;
    window.scrollTo(0, y);
  }

  /** Checks the form; returns the first problem, or null. */
  function inspProblem(tpl) {
    const items = tpl.sections.flatMap((x) => x.items);
    const missing = items.find((i) => !insp.answers[i.key]);
    if (missing) return { text: T("Answer every check: {label}", { label: pick(missing) }), key: missing.key };
    const noNote = items.find((i) => insp.answers[i.key] === "fault" && !String(insp.comments[i.key] || "").trim());
    if (noNote) return { text: T("Say what is wrong: {label}", { label: pick(noNote) }), key: noNote.key };
    if (!insp.details.result) return { text: T("Choose the overall result.") };
    if (insp.details.result === "not_fit" && !items.some((i) => insp.answers[i.key] === "fault") && !String(insp.details.notes || "").trim()) {
      return { text: T("Say why the machine is not fit for work.") };
    }
    return null;
  }

  async function submitInspection(btn) {
    const tpl = inspTpl;
    const problem = inspProblem(tpl);
    if (problem) {
      toast(problem.text, "bad");
      if (problem.key) qs(`tpi-${problem.key}`)?.scrollIntoView({ behavior: "smooth", block: "center" });
      return;
    }
    btn.disabled = true;
    insp.cid = insp.cid || uid();
    saveDraft();
    const d = insp;
    const res = await send({
      kind: "inspection",
      url: "/inspections",
      body: {
        asset_code: d.code,
        inspection_type: d.details.type,
        shift: d.details.shift,
        machine_hours: d.details.hours === "" ? null : Number(d.details.hours),
        location: d.details.location,
        overall_result: d.details.result,
        notes: d.details.notes,
        answers: d.answers,
        comments: d.comments,
        inspection_client_id: d.cid,
        lang: lang(),
      },
    });
    let queued = Boolean(res.queued);
    for (const p of d.photos) {
      const r = await send({ kind: "photo", url: "/inspections/photos", dataUrl: p.dataUrl, query: { inspection_client_id: d.cid, item_key: p.key || "" } });
      if (r.queued) queued = true;
    }
    dropDraft(d.code);
    const wo = res.data?.work_order_id;
    toast(queued ? T("Inspection saved on this phone — will send when there is signal") : wo ? T("Inspection saved — work order #{id} opened for the faults", { id: wo }) : T("Inspection saved"), "ok");
    go("inspect");
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
      img.onerror = () => reject(new Error(T("Could not read that photo")));
      img.src = url;
    });
  }

  // ------------------------------------------------------------ events
  async function onClick(e) {
    const ans = e.target.closest("[data-insp-ans],[data-insp-shift],[data-insp-result]");
    if (ans && insp) {
      if (ans.dataset.inspAns) insp.answers[ans.dataset.inspAns] = ans.dataset.v;
      if (ans.dataset.inspShift) insp.details.shift = ans.dataset.inspShift;
      if (ans.dataset.inspResult) insp.details.result = ans.dataset.inspResult;
      saveDraft();
      return render();
    }
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
      if (act === "inspectOpen") {
        const c = String(qs("tpInspCode")?.value || "").trim();
        if (c) go("inspectForm", { code: c.toUpperCase() });
        return;
      }
      if (act === "inspectStart") return go("inspectForm", { code: t.dataset.code });
      if (act === "inspPdf") return openInspectionPdf(Number(t.dataset.id));
      if (act === "inspSubmit") return submitInspection(t);
      if (act === "inspAllOk") {
        for (const i of inspTpl.sections.flatMap((x) => x.items)) if (!insp.answers[i.key]) insp.answers[i.key] = "ok";
        saveDraft();
        return render();
      }
      if (act === "inspPhotoDel") {
        insp.photos = insp.photos.filter((p) => p.id !== t.dataset.id);
        saveDraft();
        return render();
      }
      if (act === "inspDiscard") {
        sheet(`<h3>${esc(T("Discard this inspection?"))}</h3>
          <p class="tp-muted small">${esc(T("The answers and photos on this phone are deleted."))}</p>
          <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">${esc(T("Cancel"))}</button><button type="button" class="tp-btn warn grow" data-act="inspDiscardGo" data-code="${esc(t.dataset.code)}">${esc(T("Discard"))}</button></div>`);
        return;
      }
      if (act === "inspDiscardGo") {
        dropDraft(t.dataset.code);
        closeSheet();
        return go("inspect");
      }
      if (act === "openAsset") {
        const c = String(qs("tpAssetCode")?.value || "").trim();
        if (c) go("asset", { code: c.toUpperCase() });
        return;
      }
      if (act === "confirmComplete") {
        const notes = String(qs("tpDone")?.value || "").trim();
        const lead = t.dataset.lead === "1";
        if (lead && !notes) return toast(T("Say what was done first."), "bad");
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
        if (!text) return toast(T("Write what you found first."), "bad");
        const res = await send({ kind: "finding", woId: view.id, url: "/findings", body: { work_order_id: view.id, text, kind: qs("tpFindingKind").value } });
        toast(T(res.queued ? "Saved on this phone — will send later" : "Finding saved"), res.queued ? "" : "ok");
        return render();
      }
      if (act === "assetFinding") {
        const text = String(qs("tpAssetFinding")?.value || "").trim();
        if (!text) return toast(T("Write what you found first."), "bad");
        const res = await send({ kind: "finding", url: "/findings", body: { asset_id: Number(t.dataset.assetId), text } });
        toast(T(res.queued ? "Saved on this phone — will send later" : "Finding saved"), res.queued ? "" : "ok");
        return render();
      }
      if (act === "breakdownForm") {
        sheet(`<h3>${esc(T("Report breakdown on {code}", { code: t.dataset.code }))}</h3>
          <label class="tp-label">${esc(T("What failed"))}</label><input id="tpBdComp" class="tp-input" placeholder="${esc(T("e.g. Hydraulics, engine, tyre"))}" />
          <label class="tp-label">${esc(T("What happened"))}</label><textarea id="tpBdDesc" class="tp-input" rows="3"></textarea>
          <label class="tp-check"><input id="tpBdCrit" type="checkbox" /> ${esc(T("Machine cannot work (critical)"))}</label>
          <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">${esc(T("Cancel"))}</button><button type="button" class="tp-btn warn grow" data-act="breakdownSend" data-code="${esc(t.dataset.code)}">${esc(T("Report breakdown"))}</button></div>`);
        return;
      }
      if (act === "breakdownSend") {
        const description = String(qs("tpBdDesc")?.value || "").trim();
        if (!description) return toast(T("Describe what happened."), "bad");
        t.disabled = true;
        const res = await send({ kind: "breakdown", url: "/breakdowns", body: { asset_code: t.dataset.code, component: qs("tpBdComp").value, description, critical: qs("tpBdCrit").checked } });
        closeSheet();
        toast(res.queued ? T("Breakdown saved on this phone — will send later") : res.data?.work_order_id ? T("Breakdown reported — WO #{id}", { id: res.data.work_order_id }) : T("Breakdown reported"), res.queued ? "" : "ok");
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
        if (!body.part_code && !body.part_name) return toast(T("Say which part is needed."), "bad");
        t.disabled = true;
        const res = await send({ kind: "parts", woId: view.id, url: `/workorders/${view.id}/parts-request`, body });
        toast(T(res.queued ? "Request saved on this phone — will send later" : "Sent to stores"), res.queued ? "" : "ok");
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
        toast(T("Helper added"), "ok");
        return render();
      }
      if (act === "helperDel") {
        await A.fetchJson(api(`/workorders/${view.id}/helpers/${encodeURIComponent(t.dataset.user)}`), { method: "DELETE" });
        return render();
      }
      if (act === "shiftStart") {
        const res = await send({ kind: "shift", url: "/shift/start", body: {} });
        toast(T(res.queued ? "Shift start saved on this phone" : "Shift started"), "ok");
        return render();
      }
      if (act === "shiftSubmit") {
        const s = cacheGet("shift")?.data?.shift;
        sheet(`<h3>${esc(T("Submit shift report"))}</h3>
          <p>${esc(T("{h} on jobs. Jobs still running will be paused, and your job time goes to the mechanics timesheet.", { h: hrs(s?.hours || 0) }))}</p>
          ${s?.other_hours ? `<p>${esc(T("{h} other time goes to the timesheet under WORKSHOP.", { h: hrs(s.other_hours) }))}</p>` : ""}
          <p class="tp-muted small">${esc(T("After submitting, the report is locked and your foreman can see it."))}</p>
          <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">${esc(T("Cancel"))}</button><button type="button" class="tp-btn ok grow" data-act="shiftSubmitGo">${esc(T("Submit"))}</button></div>`);
        return;
      }
      if (act === "shiftSubmitGo") {
        const body = { ...read("ironlog-tech-shift-draft", {}) };
        document.querySelectorAll("[data-shift]").forEach((el) => (body[el.dataset.shift] = el.value));
        t.disabled = true;
        const res = await send({ kind: "shift", url: "/shift/submit", body });
        try { localStorage.removeItem("ironlog-tech-shift-draft"); } catch { /* ignore */ }
        closeSheet();
        toast(res.queued ? T("Report saved on this phone — will send when there is signal") : T("Shift submitted · {h}", { h: hrs(res.data?.hours || 0) }), "ok");
        return go("today");
      }
      if (act === "teamShifts") return renderTeam();
      if (act === "otherForm") {
        sheet(`<h3>${esc(T("Other time"))}</h3>
          <p class="tp-muted small">${esc(T("Time that is not on a machine. A running job is paused; starting a job stops this timer."))}</p>
          <div class="tp-reasons">${Object.keys(OTHER_REASONS).map((k) => `<label class="tp-reason"><input type="radio" name="tpReason" value="${k}" /> ${esc(reasonText(k))}</label>`).join("")}</div>
          <input id="tpOtherNote" class="tp-input" placeholder="${esc(T("Note (required for Other)"))}" />
          <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">${esc(T("Cancel"))}</button><button type="button" class="tp-btn primary grow" data-act="otherGo">${esc(T("Start timer"))}</button></div>`);
        return;
      }
      if (act === "otherGo") {
        const reason = document.querySelector("input[name=tpReason]:checked")?.value;
        const note = String(qs("tpOtherNote")?.value || "").trim();
        if (!reason) return toast(T("Choose what the time is for."), "bad");
        if (reason === "other" && !note) return toast(T("Say what the other work was."), "bad");
        t.disabled = true;
        const res = await send({ kind: "other", url: "/other/start", body: { reason, note } });
        closeSheet();
        toast(res.queued ? T("Saved on this phone — will send later") : T("Timer started: {r}", { r: reasonText(reason) }), res.queued ? "" : "ok");
        return render();
      }
      if (act === "otherStop") {
        const res = await send({ kind: "other", url: "/other/stop", body: {} });
        toast(res.queued ? T("Saved on this phone — will send later") : T("Timer stopped"), "ok");
        return render();
      }
      if (act === "unplannedForm") {
        sheet(`<h3>${esc(T("Log work on {code}", { code: t.dataset.code }))}</h3>
          <label class="tp-label">${esc(T("Part of the machine (optional)"))}</label><input id="tpUpComp" class="tp-input" placeholder="${esc(T("e.g. Mirrors, greasing, lights"))}" />
          <label class="tp-label">${esc(T("What are you doing?"))}</label><textarea id="tpUpDesc" class="tp-input" rows="3" placeholder="${esc(T("e.g. Adjusted left mirror for the operator"))}"></textarea>
          <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">${esc(T("Cancel"))}</button><button type="button" class="tp-btn primary grow" data-act="unplannedGo" data-code="${esc(t.dataset.code)}">${esc(T("Start work"))}</button></div>`);
        return;
      }
      if (act === "unplannedGo") {
        const description = String(qs("tpUpDesc")?.value || "").trim();
        if (!description) return toast(T("Say what work you are doing."), "bad");
        t.disabled = true;
        let data;
        try {
          data = await A.fetchJson(api("/unplanned"), { method: "POST", body: JSON.stringify({ asset_code: t.dataset.code, component: qs("tpUpComp").value, description, client_event_id: uid() }) });
        } catch (err) {
          t.disabled = false;
          return toast(isNetworkError(err) ? T("Opening a work order needs signal. Try again when you have signal.") : err.message || String(err), "bad");
        }
        closeSheet();
        toast(T("Work order #{id} opened — your time is running", { id: data.work_order_id }), "ok");
        return go("job", { id: Number(data.work_order_id) });
      }
      if (act === "claim") {
        sheet(`<h3>${esc(T("Take this job?"))}</h3>
          <p><b>${esc(t.dataset.label)}</b></p>
          <p class="tp-muted small">${esc(T("It will be assigned to you. You can start it straight away."))}</p>
          <div class="tp-row"><button type="button" class="tp-btn" data-act="closeSheet">${esc(T("Cancel"))}</button><button type="button" class="tp-btn primary grow" data-act="claimGo" data-wo="${t.dataset.wo}">${esc(T("Take this job"))}</button></div>`);
        return;
      }
      if (act === "claimGo") {
        t.disabled = true;
        try {
          await A.fetchJson(api(`/workorders/${t.dataset.wo}/claim`), { method: "POST", body: "{}" });
        } catch (err) {
          closeSheet();
          toast(isNetworkError(err) ? T("Taking a job needs signal. Try again when you have signal.") : err.status === 409 ? T("Someone already took this job.") : err.message || String(err), "bad");
          return render();
        }
        closeSheet();
        toast(T("The job is yours"), "ok");
        return go("job", { id: Number(t.dataset.wo) });
      }
      if (act === "ask") {
        const q = String(qs("tpAsk")?.value || "").trim();
        if (!q) return;
        const d = cached;
        const out = qs("tpAnswer");
        out.innerHTML = `<div class="tp-muted">${esc(T("Borris is thinking…"))}</div>`;
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
          lang() === "pt" ? "Answer in Portuguese (Mozambique)." : "",
        ].filter(Boolean).join(" ");
        try {
          const ans = await A.fetchJson(`${A.API}/api/ironmind/ask`, { method: "POST", body: JSON.stringify({ question: q, asset_code: d.asset.code, context_notes: context }) });
          const sources = (ans.sources || []).slice(0, 4).map((s) => esc(s.title || s.document_title || s.name || "")).filter(Boolean);
          out.innerHTML = `<div class="tp-answer">${esc(ans.short_answer || ans.answer || T("No answer."))}</div>${sources.length ? `<div class="tp-muted small">${esc(T("From:"))} ${sources.join(", ")}</div>` : ""}`;
        } catch (err) {
          out.innerHTML = `<div class="tp-error">${esc(isNetworkError(err) ? T("Borris needs signal.") : err.message || err)}</div>`;
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
    if (insp && (e.target.dataset.inspNote || e.target.dataset.inspD)) {
      if (e.target.dataset.inspNote) insp.comments[e.target.dataset.inspNote] = e.target.value;
      else insp.details[e.target.dataset.inspD] = e.target.value;
      return saveDraft();
    }
    if (e.target.id === "tpPartQ") {
      clearTimeout(partTimer);
      const q = e.target.value.trim();
      partTimer = setTimeout(async () => {
        const host = qs("tpPartHits");
        if (!host) return;
        if (q.length < 2) return (host.innerHTML = "");
        try {
          const d = await A.fetchJson(api(`/parts/search?q=${encodeURIComponent(q)}`));
          host.innerHTML = d.rows.map((p) => `<button type="button" class="tp-hit" data-act="pickPart" data-code="${esc(p.part_code)}" data-name="${esc(p.part_name)}"><b>${esc(p.part_code)}</b> ${esc(p.part_name)} <span class="tp-muted small">${esc(T("{n} in stores", { n: p.on_hand }))}${p.bin ? ` · ${esc(p.bin)}` : ""}</span></button>`).join("") || `<div class="tp-muted small">${esc(T("Nothing in stores matches. Type the part below."))}</div>`;
        } catch {
          host.innerHTML = `<div class="tp-muted small">${esc(T("Search needs signal. Type the part below."))}</div>`;
        }
      }, 300);
    }
  }

  async function onChange(e) {
    if (insp && e.target.dataset.inspD) {
      insp.details[e.target.dataset.inspD] = e.target.value;
      return saveDraft();
    }
    if (insp && e.target.dataset.inspPhoto != null && e.target.files?.length) {
      try {
        for (const f of e.target.files) insp.photos.push({ id: uid(), key: e.target.dataset.inspPhoto, dataUrl: await compress(f) });
        saveDraft();
        render();
      } catch (err) {
        toast(err.message || String(err), "bad");
      }
      return;
    }
    if (e.target.id !== "tpPhoto" || !e.target.files?.[0]) return;
    try {
      const dataUrl = await compress(e.target.files[0]);
      const entry = { kind: "photo", woId: view.id, url: `/workorders/${view.id}/photos`, dataUrl };
      const res = await send(entry);
      toast(T(res.queued ? "Photo kept on this phone — will send later" : "Photo added"), res.queued ? "" : "ok");
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

  /** Tab bar and header text, and the EN | PT buttons. */
  function translateChrome() {
    document.documentElement.lang = lang();
    document.querySelectorAll("[data-t]").forEach((el) => (el.textContent = T(el.dataset.t)));
    document.querySelectorAll("[data-lang]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.lang === lang())));
  }

  function setLang(l) {
    try { localStorage.setItem(LANG_KEY, l); } catch { /* ignore */ }
    translateChrome();
    if (root) render();
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
    else if (p.get("inspect")) view = { name: "inspectForm", code: p.get("inspect").toUpperCase() };
    else if (["shift", "week", "inspect"].includes(p.get("tab"))) view = { name: p.get("tab") };
    registerWorker();
    render();
    syncQueue();
  }

  // The language switch works on the sign-in screen too.
  document.querySelectorAll("[data-lang]").forEach((b) => b.addEventListener("click", () => setLang(b.dataset.lang)));
  translateChrome();

  window.TechPortal = { start, syncQueue, T, lang, translateChrome };
})();

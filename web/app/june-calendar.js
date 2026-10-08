// IRONLOG/web/app/june-calendar.js — Outlook-style calendar next to June on My Work.
// Part of the main app; index.html loads these files in order and they share one global scope.
//
// One calendar for two sources: the private calendar link (ICS, e.g. Outlook)
// and June's Ironlog schedule. Day, work week, week and month views; events
// sit on an hour grid like Outlook. june.js loads the events for the range
// shown and hands them over with JuneCalendar.setSource().

const JuneCalendar = (() => {
  const VIEW_KEY = "ironlog_june_cal_view";
  const HIDE_KEY = "ironlog_june_cal_hidden";
  const HOUR_PX = 44;
  const SOURCES = {
    ics: { label: "My calendar", cls: "ics" },
    ironlog: { label: "Ironlog schedule", cls: "ironlog" },
  };
  const DOW = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
  let view = "week";
  let anchor = new Date();
  let events = { ics: [], ironlog: [] };
  let errors = { ics: "", ironlog: "" };
  let icsConnected = false;
  let hidden = new Set();
  let onRange = null;
  let nowTimer = null;

  try { view = localStorage.getItem(VIEW_KEY) || "week"; } catch { /* storage off */ }
  try { hidden = new Set(JSON.parse(localStorage.getItem(HIDE_KEY) || "[]")); } catch { hidden = new Set(); }

  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const pad = (n) => String(n).padStart(2, "0");
  const ymd = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  const addDays = (d, n) => { const x = new Date(d); x.setDate(x.getDate() + n); return x; };
  const startOfDay = (d) => { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; };
  const mondayOf = (d) => { const x = startOfDay(d); const dow = (x.getDay() + 6) % 7; return addDays(x, -dow); };
  const hm = (d) => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
  const fromYmd = (s) => { const [y, m, d] = String(s).split("-").map(Number); return new Date(y, m - 1, d); };

  /** The days shown for the current view. */
  function days() {
    if (view === "day") return [startOfDay(anchor)];
    if (view === "workweek") return Array.from({ length: 5 }, (_, i) => addDays(mondayOf(anchor), i));
    if (view === "month") {
      const first = new Date(anchor.getFullYear(), anchor.getMonth(), 1);
      const start = mondayOf(first);
      return Array.from({ length: 42 }, (_, i) => addDays(start, i));
    }
    return Array.from({ length: 7 }, (_, i) => addDays(mondayOf(anchor), i));
  }

  /** The date range june.js should load: the visible days. */
  function range() {
    const list = days();
    return { from: ymd(list[0]), to: ymd(list[list.length - 1]) };
  }

  /** Events from both sources in one shape: { src, title, start, end, allDay, … }. */
  function normalised() {
    const out = [];
    for (const e of events.ics || []) {
      if (e.all_day) {
        // All-day ICS dates are calendar dates (UTC midnight to UTC midnight).
        const startStr = e.start_date || String(e.start_at || "").slice(0, 10);
        const endStr = e.end_at ? new Date(Date.parse(e.end_at) - 1).toISOString().slice(0, 10) : startStr;
        out.push({ src: "ics", title: e.summary || "(No title)", allDay: true, start: fromYmd(startStr), end: fromYmd(endStr < startStr ? startStr : endStr), location: e.location || "", raw: e });
      } else {
        const start = new Date(e.start_at);
        const end = e.end_at ? new Date(e.end_at) : new Date(start.getTime() + 3_600_000);
        if (Number.isNaN(start.getTime())) continue;
        out.push({ src: "ics", title: e.summary || "(No title)", allDay: false, start, end: end > start ? end : new Date(start.getTime() + 1_800_000), location: e.location || "", raw: e });
      }
    }
    for (const e of events.ironlog || []) {
      const date = e.event_date || e.start_date;
      if (!date) continue;
      const detail = [String(e.category || "").replace(/_/g, " "), e.asset_code ? `Asset ${e.asset_code}` : "", e.work_order_id ? `WO #${e.work_order_id}` : ""].filter(Boolean).join(" · ");
      if (e.all_day || !e.start_time) {
        out.push({ src: "ironlog", title: e.title || "(No title)", allDay: true, start: fromYmd(date), end: fromYmd(date), location: detail, notes: e.notes || "", raw: e });
      } else {
        const start = new Date(`${date}T${e.start_time}`);
        let end = e.end_time ? new Date(`${date}T${e.end_time}`) : new Date(start.getTime() + 3_600_000);
        if (!(end > start)) end = new Date(start.getTime() + 1_800_000);
        out.push({ src: "ironlog", title: e.title || "(No title)", allDay: false, start, end, location: detail, notes: e.notes || "", raw: e });
      }
    }
    return out.filter((e) => !hidden.has(e.src));
  }

  function titleText() {
    const list = days();
    const fmt = (d, o) => d.toLocaleDateString(undefined, o);
    if (view === "day") return fmt(list[0], { weekday: "long", day: "numeric", month: "long", year: "numeric" });
    if (view === "month") return fmt(anchor, { month: "long", year: "numeric" });
    const a = list[0];
    const b = list[list.length - 1];
    if (a.getMonth() === b.getMonth()) return `${a.getDate()}–${b.getDate()} ${fmt(b, { month: "long", year: "numeric" })}`;
    return `${fmt(a, { day: "numeric", month: "short" })} – ${fmt(b, { day: "numeric", month: "short", year: "numeric" })}`;
  }

  /** Side-by-side columns for overlapping timed events in one day. */
  function layoutDay(list) {
    const sorted = [...list].sort((a, b) => a.start - b.start || b.end - a.end);
    const placed = [];
    let cluster = [];
    let clusterEnd = 0;
    const flush = () => {
      const cols = Math.max(1, ...cluster.map((p) => p.col + 1));
      cluster.forEach((p) => { p.cols = cols; });
      cluster = [];
    };
    for (const e of sorted) {
      if (cluster.length && e.start >= clusterEnd) flush();
      const used = new Set(cluster.filter((p) => p.e.end > e.start).map((p) => p.col));
      let col = 0;
      while (used.has(col)) col += 1;
      const p = { e, col, cols: 1 };
      cluster.push(p);
      placed.push(p);
      clusterEnd = Math.max(clusterEnd, e.end.getTime());
    }
    flush();
    return placed;
  }

  function renderLegend() {
    const host = $("jcLegend");
    if (!host) return;
    host.innerHTML = Object.entries(SOURCES).map(([k, s]) => {
      if (k === "ics" && !icsConnected) return "";
      return `<label class="jc-src ${s.cls}"><input type="checkbox" data-jc-src="${k}" ${hidden.has(k) ? "" : "checked"} /> <span class="jc-swatch"></span>${esc(k === "ics" ? (SOURCES.ics.name || s.label) : s.label)}</label>`;
    }).join("") + Object.entries(errors).filter(([, v]) => v).map(([k, v]) => `<span class="jc-err" title="${esc(v)}">⚠ ${esc(SOURCES[k].label)} could not refresh</span>`).join("");
  }

  function render() {
    const body = $("jcBody");
    if (!body) return;
    if ($("jcTitle")) $("jcTitle").textContent = titleText();
    document.querySelectorAll("[data-jc-view]").forEach((b) => b.classList.toggle("on", b.dataset.jcView === view));
    renderLegend();
    const list = normalised();
    if (view === "month") return renderMonth(body, list);
    renderGrid(body, list);
  }

  function renderGrid(body, list) {
    const ds = days();
    const today = ymd(new Date());
    const timed = list.filter((e) => !e.allDay);
    // Working hours by default; widen to fit the events shown.
    let first = 7;
    let last = 18;
    for (const e of timed) {
      if (ds.some((d) => ymd(e.start) === ymd(d))) {
        first = Math.min(first, e.start.getHours());
        last = Math.max(last, e.end.getHours() + (e.end.getMinutes() ? 1 : 0));
      }
    }
    first = Math.max(0, first);
    last = Math.min(24, Math.max(last, first + 1));
    const hours = Array.from({ length: last - first }, (_, i) => first + i);
    const cols = `56px repeat(${ds.length}, minmax(0, 1fr))`;
    const allDay = ds.map((d) => list.filter((e) => e.allDay && startOfDay(e.start) <= d && startOfDay(e.end) >= d));
    body.innerHTML = `
      <div class="jc-grid-head" style="grid-template-columns:${cols}">
        <div></div>
        ${ds.map((d) => `<div class="jc-dayhead ${ymd(d) === today ? "today" : ""}"><span>${DOW[(d.getDay() + 6) % 7]}</span><b>${d.getDate()}</b></div>`).join("")}
      </div>
      ${allDay.some((x) => x.length) ? `
        <div class="jc-allday" style="grid-template-columns:${cols}">
          <div class="jc-gutter-label">All day</div>
          ${allDay.map((x) => `<div class="jc-allday-cell">${x.map((e) => chip(e)).join("")}</div>`).join("")}
        </div>` : ""}
      <div class="jc-scroll" id="jcScroll">
        <div class="jc-grid" style="grid-template-columns:${cols}; height:${hours.length * HOUR_PX}px">
          <div class="jc-gutter">${hours.map((h) => `<div class="jc-hour" style="height:${HOUR_PX}px"><span>${pad(h)}:00</span></div>`).join("")}</div>
          ${ds.map((d) => {
            const dayEvents = timed.filter((e) => ymd(e.start) === ymd(d));
            const placed = layoutDay(dayEvents);
            const isToday = ymd(d) === today;
            const now = new Date();
            const nowTop = ((now.getHours() - first) * 60 + now.getMinutes()) * (HOUR_PX / 60);
            return `<div class="jc-col ${isToday ? "today" : ""}">
              ${hours.map(() => `<div class="jc-slot" style="height:${HOUR_PX}px"></div>`).join("")}
              ${placed.map((p) => {
                const top = ((p.e.start.getHours() - first) * 60 + p.e.start.getMinutes()) * (HOUR_PX / 60);
                const height = Math.max(20, ((p.e.end - p.e.start) / 60000) * (HOUR_PX / 60) - 2);
                const width = 100 / p.cols;
                return `<button type="button" class="jc-ev ${SOURCES[p.e.src].cls}" data-jc-ev="${esc(eventKey(p.e))}" style="top:${top}px;height:${height}px;left:calc(${p.col * width}% + 2px);width:calc(${width}% - 4px)">
                  <b>${esc(p.e.title)}</b>${height > 30 ? `<span>${hm(p.e.start)}–${hm(p.e.end)}${p.e.location ? ` · ${esc(p.e.location)}` : ""}</span>` : ""}
                </button>`;
              }).join("")}
              ${isToday && nowTop >= 0 && nowTop <= hours.length * HOUR_PX ? `<div class="jc-now" style="top:${nowTop}px"></div>` : ""}
            </div>`;
          }).join("")}
        </div>
      </div>`;
    // Start the view at the morning (or now), like Outlook.
    const scroll = $("jcScroll");
    if (scroll) {
      const now = new Date();
      const target = ds.some((d) => ymd(d) === today) ? Math.max(0, now.getHours() - first - 1) : Math.max(0, 7 - first);
      scroll.scrollTop = target * HOUR_PX;
    }
  }

  function chip(e) {
    return `<button type="button" class="jc-chip ${SOURCES[e.src].cls}" data-jc-ev="${esc(eventKey(e))}">${e.allDay ? "" : `<span>${hm(e.start)}</span> `}${esc(e.title)}</button>`;
  }

  function renderMonth(body, list) {
    const ds = days();
    const month = anchor.getMonth();
    const today = ymd(new Date());
    body.innerHTML = `
      <div class="jc-month-head">${DOW.map((d) => `<div>${d}</div>`).join("")}</div>
      <div class="jc-month">
        ${ds.map((d) => {
          const dayEvents = list.filter((e) => (e.allDay ? startOfDay(e.start) <= d && startOfDay(e.end) >= d : ymd(e.start) === ymd(d))).sort((a, b) => Number(b.allDay) - Number(a.allDay) || a.start - b.start);
          const more = dayEvents.length - 3;
          return `<div class="jc-mday ${d.getMonth() !== month ? "other" : ""} ${ymd(d) === today ? "today" : ""}">
            <button type="button" class="jc-mdate" data-jc-day="${ymd(d)}">${d.getDate()}</button>
            ${dayEvents.slice(0, 3).map((e) => chip(e)).join("")}
            ${more > 0 ? `<button type="button" class="jc-more" data-jc-day="${ymd(d)}">+${more} more</button>` : ""}
          </div>`;
        }).join("")}
      </div>`;
  }

  const eventKey = (e) => `${e.src}|${e.start.getTime()}|${e.title}`;

  function showPopover(btn) {
    const pop = $("jcPop");
    const e = normalised().find((x) => eventKey(x) === btn.dataset.jcEv);
    if (!pop || !e) return;
    const when = e.allDay
      ? `${e.start.toLocaleDateString(undefined, { weekday: "long", day: "numeric", month: "long" })} · All day`
      : `${e.start.toLocaleDateString(undefined, { weekday: "long", day: "numeric", month: "long" })} · ${hm(e.start)}–${hm(e.end)}`;
    pop.innerHTML = `
      <div class="jc-pop-head ${SOURCES[e.src].cls}"><b>${esc(e.title)}</b><button type="button" class="jc-pop-x" data-jc-close aria-label="Close">✕</button></div>
      <div class="jc-pop-row">🕒 ${esc(when)}</div>
      ${e.location ? `<div class="jc-pop-row">${e.src === "ics" ? "📍" : "🔧"} ${esc(e.location)}</div>` : ""}
      ${e.notes ? `<div class="jc-pop-row">${esc(e.notes)}</div>` : ""}
      <div class="jc-pop-src">${esc(e.src === "ics" ? (SOURCES.ics.name || "My calendar") + " (read-only)" : "Ironlog schedule · ask June to move or cancel it")}</div>`;
    const card = $("juneCalendarCard").getBoundingClientRect();
    const r = btn.getBoundingClientRect();
    pop.hidden = false;
    const w = pop.offsetWidth;
    let left = r.right - card.left + 8;
    if (left + w > card.width - 8) left = Math.max(8, r.left - card.left - w - 8);
    pop.style.left = `${left}px`;
    pop.style.top = `${Math.max(8, Math.min(r.top - card.top, card.height - pop.offsetHeight - 8))}px`;
  }

  function go(step) {
    if (view === "day") anchor = addDays(anchor, step);
    else if (view === "month") anchor = new Date(anchor.getFullYear(), anchor.getMonth() + step, 1);
    else anchor = addDays(anchor, 7 * step);
    changed();
  }
  function changed() {
    $("jcPop") && ($("jcPop").hidden = true);
    render();
    onRange?.(range());
  }

  function wire() {
    const card = $("juneCalendarCard");
    if (!card || card.dataset.jcWired) return;
    card.dataset.jcWired = "1";
    card.addEventListener("click", (e) => {
      const t = e.target;
      if (t.closest("[data-jc-today]")) { anchor = new Date(); return changed(); }
      if (t.closest("[data-jc-prev]")) return go(-1);
      if (t.closest("[data-jc-next]")) return go(1);
      const v = t.closest("[data-jc-view]");
      if (v) {
        view = v.dataset.jcView;
        try { localStorage.setItem(VIEW_KEY, view); } catch { /* storage off */ }
        return changed();
      }
      const day = t.closest("[data-jc-day]");
      if (day) { anchor = fromYmd(day.dataset.jcDay); view = "day"; return changed(); }
      const ev = t.closest("[data-jc-ev]");
      if (ev) return showPopover(ev);
      if (t.closest("[data-jc-close]") || !t.closest("#jcPop")) { if ($("jcPop")) $("jcPop").hidden = true; }
    });
    card.addEventListener("change", (e) => {
      const src = e.target.closest("[data-jc-src]");
      if (!src) return;
      if (src.checked) hidden.delete(src.dataset.jcSrc);
      else hidden.add(src.dataset.jcSrc);
      try { localStorage.setItem(HIDE_KEY, JSON.stringify([...hidden])); } catch { /* storage off */ }
      render();
    });
    clearInterval(nowTimer);
    nowTimer = setInterval(() => { if (!document.hidden && view !== "month") render(); }, 5 * 60 * 1000);
  }

  return {
    start(loader) {
      onRange = loader;
      wire();
      render();
      onRange?.(range());
    },
    range,
    render,
    setSource(src, list, error = "") {
      events[src] = Array.isArray(list) ? list : [];
      errors[src] = error || "";
      render();
    },
    setIcs({ connected, name }) {
      icsConnected = Boolean(connected);
      SOURCES.ics.name = name || "My calendar";
      if (!icsConnected) { events.ics = []; errors.ics = ""; }
      render();
    },
  };
})();
window.JuneCalendar = JuneCalendar;

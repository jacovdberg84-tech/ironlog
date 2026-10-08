// IRONLOG/web/app/june.js — June, the admin-only GPT-Live task assistant.
// The OpenAI project key is never exposed to this page. WebRTC negotiation is
// proxied through /api/june/live/session and all private tool calls are sent
// to Ironlog's authenticated June gateway.
let junePeerConnection = null;
let juneDataChannel = null;
let juneLocalStream = null;
let juneAudio = null;
let juneProcessedCalls = new Set();
let juneIcsCalendar = null;
let juneAvatarConfig = null;
let juneRemoteAudioStream = null;
let juneLemonSlice = null;
let juneLiveGeneration = 0;
let juneLiveKitLoadPromise = null;

function juneLemonSliceEnabled() {
  return Boolean(juneAvatarConfig?.configured);
}

function juneEnsureLiveKitClient() {
  if (window.LivekitClient?.Room && window.LivekitClient?.RoomEvent) return Promise.resolve(window.LivekitClient);
  if (juneLiveKitLoadPromise) return juneLiveKitLoadPromise;
  juneLiveKitLoadPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/livekit-client@2.22.3/dist/livekit-client.umd.min.js";
    script.async = true;
    script.dataset.juneLivekit = "true";
    script.onload = () => window.LivekitClient?.Room
      ? resolve(window.LivekitClient)
      : reject(new Error("June's visual library did not initialise."));
    script.onerror = () => reject(new Error("June could not load her visual library."));
    document.head.appendChild(script);
  }).catch((error) => {
    juneLiveKitLoadPromise = null;
    throw error;
  });
  return juneLiveKitLoadPromise;
}

function juneLemonSliceClearVideo() {
  const video = qs("juneAvatarVideo");
  const avatar = qs("juneAvatar");
  if (video) {
    try { video.pause(); } catch {}
    video.srcObject = null;
    video.hidden = true;
  }
  if (avatar) delete avatar.dataset.renderer;
}

function juneLemonSliceFail(error) {
  // The animated layer is intentionally best-effort. A provider or network
  // issue must not break the existing GPT-Live conversation.
  if (!juneLemonSlice?.warned) {
    console.warn("June visual paused; voice remains available.", error);
    juneLemonSlice = { ...juneLemonSlice, warned: true, enabled: false };
    juneAvatarNote(`Animated avatar off: ${error?.message || String(error)} Voice still works.`);
  }
  juneLemonSliceDirectVoice(juneLemonSlice);
  juneLemonSliceClearVideo();
}

/** A short line under June saying why her LemonSlice avatar is not showing. */
function juneAvatarNote(text) {
  const note = qs("juneAvatarNote");
  if (!note) return;
  note.textContent = String(text || "");
  note.hidden = !text;
}

function juneLemonSliceBase64(bytes) {
  const view = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let text = "";
  for (let i = 0; i < view.length; i += 0x8000) {
    text += String.fromCharCode(...view.subarray(i, i + 0x8000));
  }
  return btoa(text);
}

function juneLemonSliceResample(input, sourceRate) {
  const rate = Number(sourceRate || 48_000);
  const outputLength = Math.max(1, Math.round(input.length * 16_000 / rate));
  const output = new Int16Array(outputLength);
  for (let index = 0; index < outputLength; index += 1) {
    const start = Math.floor(index * rate / 16_000);
    const end = Math.min(input.length, Math.max(start + 1, Math.floor((index + 1) * rate / 16_000)));
    let sum = 0;
    for (let sample = start; sample < end; sample += 1) sum += input[sample] || 0;
    output[index] = Math.max(-1, Math.min(1, sum / Math.max(1, end - start))) * 0x7fff;
  }
  return output;
}

function juneLemonSliceQueue(path, body = {}) {
  const state = juneLemonSlice;
  if (!state?.enabled || !state?.id) return Promise.resolve();
  const id = encodeURIComponent(state.id);
  state.queue = state.queue.then(async () => {
    if (!juneLemonSlice?.enabled || juneLemonSlice.id !== state.id) return;
    const response = await fetch(`${API}/api/june/avatar/session/${id}/${path}`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify(body),
    });
    if (!response.ok) throw new Error(`June visual stream returned HTTP ${response.status}.`);
  }).catch((error) => juneLemonSliceFail(error));
  return state.queue;
}

function juneLemonSliceFlushAudio({ force = false } = {}) {
  const state = juneLemonSlice;
  if (!state?.enabled || !state.pending.length) return;
  if (state.audioInFlight && !force) return;
  const length = state.pending.reduce((total, chunk) => total + chunk.length, 0);
  const combined = new Int16Array(length);
  let offset = 0;
  for (const chunk of state.pending) {
    combined.set(chunk, offset);
    offset += chunk.length;
  }
  state.pending = [];
  state.audioInFlight = true;
  juneLemonSliceQueue("audio", { audio: juneLemonSliceBase64(combined) }).finally(() => {
    if (juneLemonSlice !== state) return;
    state.audioInFlight = false;
    const waiting = state.pending.reduce((total, chunk) => total + chunk.length, 0);
    if (waiting >= 800) juneLemonSliceFlushAudio();
  });
}

function juneLemonSliceStartAudio(stream) {
  const state = juneLemonSlice;
  if (!state?.enabled || state.processor || !stream) return;
  try {
    const context = new AudioContext();
    const source = context.createMediaStreamSource(stream);
    const processor = context.createScriptProcessor(2048, 1, 1);
    const silence = context.createGain();
    silence.gain.value = 0;
    processor.onaudioprocess = (event) => {
      const active = juneLemonSlice;
      if (!active?.enabled || active.id !== state.id) return;
      const pcm = juneLemonSliceResample(event.inputBuffer.getChannelData(0), event.inputBuffer.sampleRate);
      active.pending.push(pcm);
      const pendingSamples = active.pending.reduce((total, chunk) => total + chunk.length, 0);
      // ~100 ms chunks, as LemonSlice recommends; while one is on its way the
      // next ones are gathered and sent together, so the backlog never grows.
      if (pendingSamples >= 1_600) juneLemonSliceFlushAudio();
    };
    source.connect(processor);
    processor.connect(silence);
    silence.connect(context.destination);
    state.audioContext = context;
    state.source = source;
    state.processor = processor;
    state.silence = silence;
    context.resume().catch(() => {});
  } catch (error) {
    juneLemonSliceFail(error);
  }
}

function juneLemonSliceCommitResponse() {
  if (!juneLemonSlice?.enabled) return;
  juneLemonSliceFlushAudio({ force: true });
  juneLemonSliceQueue("end-turn");
}

function juneLemonSliceInterrupt() {
  if (!juneLemonSlice?.enabled) return;
  juneLemonSlice.pending = [];
  juneLemonSliceQueue("interrupt");
}

async function juneStartLemonSliceAvatar(generation) {
  if (!juneLemonSliceEnabled()) return false;
  try {
    const livekit = await juneEnsureLiveKitClient();
    const data = await juneApiJson(`${API}/api/june/avatar/session`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: "{}",
    });
    const session = data?.session;
    if (!session?.id || !session?.livekit_url || !session?.viewer_token) throw new Error("June's visual session response was incomplete.");
    // June may have been stopped/restarted while LemonSlice was preparing.
    // Close that now-orphaned server session instead of attaching it to a
    // newer conversation.
    if (generation !== juneLiveGeneration) {
      fetch(`${API}/api/june/avatar/session/${encodeURIComponent(session.id)}/stop`, {
        method: "POST",
        headers: authHeaders({ "Content-Type": "application/json" }),
        body: "{}",
        keepalive: true,
      }).catch(() => {});
      return false;
    }
    const room = new livekit.Room({ adaptiveStream: true, dynacast: false });
    juneLemonSlice = {
      id: String(session.id), room, enabled: true, warned: false, queue: Promise.resolve(), pending: [],
      audioContext: null, source: null, processor: null, silence: null,
    };
    room.on(livekit.RoomEvent.TrackSubscribed, (track) => {
      const kind = String(track?.kind || "").toLowerCase();
      if (kind === "audio") {
        // The avatar's own audio is in step with its lips. It replaces the
        // direct voice once the video is showing (juneLemonSliceUseAvatarVoice).
        juneLemonSlice.avatarAudio = track;
        try { track.setVolume?.(0); } catch {}
        juneLemonSliceUseAvatarVoice();
        return;
      }
      if (kind !== "video") return;
      const video = qs("juneAvatarVideo");
      const avatar = qs("juneAvatar");
      if (!video || !avatar) return;
      track.attach(video);
      video.hidden = false;
      avatar.dataset.renderer = "lemonslice";
      video.play().catch(() => {});
      juneLemonSlice.videoShown = true;
      juneLemonSliceUseAvatarVoice();
    });
    room.on(livekit.RoomEvent.Disconnected, () => {
      if (juneLemonSlice?.id === session.id) juneLemonSliceFail(new Error("June's visual stream closed."));
    });
    await room.connect(session.livekit_url, session.viewer_token);
    if (juneRemoteAudioStream) juneLemonSliceStartAudio(juneRemoteAudioStream);
    return true;
  } catch (error) {
    juneLemonSliceFail(error);
    return false;
  }
}

/** Hear June through the avatar (lips and voice in step) instead of the direct stream. */
function juneLemonSliceUseAvatarVoice() {
  const state = juneLemonSlice;
  if (!state?.enabled || !state.videoShown || !state.avatarAudio || state.avatarVoiceOn) return;
  try {
    const el = state.avatarAudio.attach();
    el.hidden = true;
    document.body.appendChild(el);
    state.avatarAudioEl = el;
    state.avatarAudio.setVolume?.(1);
    el.play?.().catch(() => {});
    if (juneAudio) juneAudio.muted = true;
    state.avatarVoiceOn = true;
  } catch {
    if (juneAudio) juneAudio.muted = false;
  }
}

/** Back to the direct voice (avatar stopped or failed). */
function juneLemonSliceDirectVoice(state) {
  if (juneAudio) juneAudio.muted = false;
  try { state?.avatarAudio?.detach?.(); } catch {}
  try { state?.avatarAudioEl?.remove(); } catch {}
}

function juneStopLemonSliceAvatar() {
  const state = juneLemonSlice;
  juneLemonSlice = null;
  if (!state) {
    juneLemonSliceClearVideo();
    return;
  }
  try { state.processor?.disconnect(); } catch {}
  try { state.source?.disconnect(); } catch {}
  try { state.silence?.disconnect(); } catch {}
  try { state.audioContext?.close(); } catch {}
  try { state.room?.disconnect(); } catch {}
  juneLemonSliceDirectVoice(state);
  juneLemonSliceClearVideo();
  if (state.id) {
    fetch(`${API}/api/june/avatar/session/${encodeURIComponent(state.id)}/stop`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: "{}",
      keepalive: true,
    }).catch(() => {});
  }
}

// The portrait layer is optional visual polish. It must never be able to
// interrupt June's live voice session if a browser cannot load or analyse it.
function juneSetAvatarState(state) {
  try {
    if (typeof setJuneAvatarVisualState === "function") setJuneAvatarVisualState(state);
  } catch {}
}

function juneStartAvatarAudio(stream) {
  try {
    if (typeof startJuneAvatarAudioAnalysis === "function") startJuneAvatarAudioAnalysis(stream);
  } catch {}
}

function juneStopAvatarAudio() {
  try {
    if (typeof stopJuneAvatarAudioAnalysis === "function") stopJuneAvatarAudioAnalysis();
  } catch {}
}

function juneIsAdmin() {
  return getSessionRoles().includes("admin");
}

function juneEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function juneSetState(message, tone = "neutral") {
  const el = qs("juneLiveState");
  if (!el) return;
  el.textContent = message;
  el.dataset.tone = tone;
  const visual = tone === "live" ? "listening" : tone === "working" ? "thinking" : "idle";
  juneSetAvatarState(visual);
}

// ---- Memory: the conversation is saved so the next session can pick it up.
const juneMem = { user: "", june: "", sawSessionOutput: false, buf: [], timer: null };

function juneMemAdd(role, text) {
  const t = String(text || "").replace(/\s+/g, " ").trim();
  if (!t) return;
  juneMem.buf.push({ role, text: t });
  clearTimeout(juneMem.timer);
  juneMem.timer = setTimeout(() => juneMemSend(), 1500);
}
function juneMemFlushUser() { if (juneMem.user.trim()) juneMemAdd("user", juneMem.user); juneMem.user = ""; }
function juneMemFlushJune() { if (juneMem.june.trim()) juneMemAdd("june", juneMem.june); juneMem.june = ""; }
function juneMemSend({ final = false } = {}) {
  clearTimeout(juneMem.timer);
  if (!juneMem.buf.length) return;
  const turns = juneMem.buf.splice(0, juneMem.buf.length);
  fetch(`${API}/api/june/memory/turns`, {
    method: "POST",
    headers: authHeaders({ "Content-Type": "application/json" }),
    body: JSON.stringify({ turns }),
    keepalive: final,
  }).catch(() => {});
}
async function juneForgetMemory() {
  if (!confirm("Clear June's memory of your past conversations?")) return;
  try {
    await juneApiJson(`${API}/api/june/memory/clear`, { method: "POST", headers: authHeaders({ "Content-Type": "application/json" }), body: "{}" });
    juneSetState("June has forgotten your past conversations.", "ready");
  } catch (error) {
    juneSetState(`June could not clear her memory: ${error?.message || error}`, "warning");
  }
}

function juneAppendTranscript(speaker, content, kind = "assistant") {
  const host = qs("juneTranscript");
  const text = String(content || "").trim();
  if (!host || !text) return;
  const row = document.createElement("div");
  row.className = `june-transcript-row june-transcript-${kind}`;
  const name = document.createElement("span");
  name.className = "june-transcript-speaker";
  name.textContent = speaker;
  const message = document.createElement("div");
  message.className = "june-transcript-message";
  message.textContent = text;
  row.append(name, message);
  host.appendChild(row);
  host.scrollTop = host.scrollHeight;
}

function juneClearTranscript() {
  const host = qs("juneTranscript");
  if (host) host.innerHTML = "";
}

function juneRenderMaintenanceScheduleDownload(result) {
  const download = result?.download;
  const reportId = String(download?.report_id || "").trim();
  const host = qs("juneDraftDownloads");
  if (!host || !reportId || String(result?.type || "") !== "maintenance_schedule") return;
  const summary = result?.summary || {};
  const selected = Number(summary.selected_assets || 0);
  const overdue = Number(summary.overdue || 0);
  const horizon = Number(result?.horizon_days || 30);
  const expires = Math.max(1, Number(download?.expires_in_minutes || 15));
  const card = document.createElement("div");
  card.className = "june-draft-download";
  const detail = document.createElement("div");
  const title = document.createElement("strong");
  title.textContent = "Draft maintenance schedule ready";
  const note = document.createElement("span");
  note.textContent = `${selected} asset${selected === 1 ? "" : "s"} · ${horizon}-day view${overdue ? ` · ${overdue} overdue` : ""} · review-only · available for ${expires} minutes`;
  detail.append(title, note);
  const button = document.createElement("button");
  button.type = "button";
  button.className = "btn btn-secondary btn-sm";
  button.textContent = String(download?.label || "Download Excel schedule");
  button.addEventListener("click", async () => {
    button.disabled = true;
    const filename = String(download?.filename || "IRONLOG_June_Maintenance_Schedule.xlsx");
    try {
      await downloadAuthedFile(`${API}/api/june/maintenance-schedule/${encodeURIComponent(reportId)}.xlsx`, filename);
    } finally {
      button.disabled = false;
    }
  });
  card.append(detail, button);
  host.replaceChildren(card);
  host.hidden = false;
}

function juneCalendarEventWhen(event) {
  const date = String(event?.event_date || event?.start_date || "").trim();
  if (event?.all_day) return `${date} · All day`;
  const start = String(event?.start_time || "").trim();
  const end = String(event?.end_time || "").trim();
  return `${date}${start ? ` · ${start}` : ""}${end ? `–${end}` : ""}`;
}

function juneRenderInternalCalendarEvents(events, { error = "" } = {}) {
  window.JuneCalendar?.setSource("ironlog", events, error);
}

async function juneLoadInternalCalendar({ quiet = false } = {}) {
  const refresh = qs("juneInternalCalendarRefreshBtn");
  if (refresh) refresh.disabled = true;
  try {
    const range = window.JuneCalendar?.range();
    const query = range ? `from=${range.from}&to=${range.to}` : "days=21";
    const data = await juneApiJson(`${API}/api/june/calendar/internal/events?${query}`, { headers: authHeaders() });
    juneRenderInternalCalendarEvents(data?.calendar?.events || []);
  } catch (error) {
    juneRenderInternalCalendarEvents([], { error: `Ironlog schedule could not be refreshed: ${error?.message || String(error)}` });
    if (!quiet) juneSetState("June could not refresh the Ironlog schedule.", "warning");
  } finally {
    if (refresh) refresh.disabled = false;
  }
}

function juneHideCalendarApproval() {
  const host = qs("juneCalendarApproval");
  if (!host) return;
  host.hidden = true;
  host.replaceChildren();
}

function juneRenderCalendarApproval(result) {
  const host = qs("juneCalendarApproval");
  const approval = result?.approval || {};
  const token = String(approval?.token || "").trim();
  const event = result?.event || {};
  const action = String(result?.action || "").toLowerCase();
  if (!host || !token || !["create", "update", "cancel"].includes(action)) return;
  const card = document.createElement("div");
  card.className = "june-calendar-approval-card";
  const copy = document.createElement("div");
  const title = document.createElement("strong");
  title.textContent = action === "cancel" ? "June is ready to cancel this entry" : action === "update" ? "June is ready to update this entry" : "June is ready to add this entry";
  const detail = document.createElement("span");
  detail.textContent = `${String(event?.title || "Calendar entry")} · ${juneCalendarEventWhen(event)}${event?.asset_code ? ` · ${event.asset_code}` : ""}`;
  const note = document.createElement("small");
  note.textContent = action === "cancel" ? "Confirming removes it from the active Ironlog schedule." : "Nothing changes until you confirm.";
  copy.append(title, detail, note);
  const actions = document.createElement("div");
  actions.className = "june-calendar-approval-actions";
  const discard = document.createElement("button");
  discard.type = "button";
  discard.className = "btn btn-secondary btn-sm";
  discard.textContent = "Discard";
  const confirm = document.createElement("button");
  confirm.type = "button";
  confirm.className = action === "cancel" ? "btn btn-danger btn-sm" : "btn btn-primary btn-sm";
  confirm.textContent = action === "cancel" ? "Confirm cancellation" : "Confirm change";
  discard.addEventListener("click", async () => {
    discard.disabled = true;
    confirm.disabled = true;
    try {
      await juneApiJson(`${API}/api/june/calendar/internal/approvals/${encodeURIComponent(token)}/discard`, {
        method: "POST",
        headers: authHeaders({ "Content-Type": "application/json" }),
        body: "{}",
      });
      juneHideCalendarApproval();
      juneSetState("Calendar change discarded. Sensible restraint for once.", "ready");
    } catch (error) {
      confirm.disabled = false;
      juneSetState(`Calendar change could not be discarded: ${error?.message || String(error)}`, "warning");
    }
  });
  confirm.addEventListener("click", async () => {
    discard.disabled = true;
    confirm.disabled = true;
    juneSetState("June is applying your confirmed calendar change…", "working");
    try {
      const data = await juneApiJson(`${API}/api/june/calendar/internal/approvals/${encodeURIComponent(token)}/confirm`, {
        method: "POST",
        headers: authHeaders({ "Content-Type": "application/json" }),
        body: "{}",
      });
      juneHideCalendarApproval();
      await juneLoadInternalCalendar({ quiet: true });
      const message = String(data?.message || "Ironlog schedule updated.");
      juneAppendTranscript("June", message, "assistant");
      juneSetState(message, "ready");
    } catch (error) {
      confirm.disabled = false;
      discard.disabled = false;
      juneSetState(`Calendar change could not be applied: ${error?.message || String(error)}`, "warning");
    }
  });
  actions.append(discard, confirm);
  card.append(copy, actions);
  host.replaceChildren(card);
  host.hidden = false;
}

function juneCalendarToolOutputForLive(result) {
  const approval = result?.approval || {};
  const { token: _token, ...approvalForLive } = approval;
  return {
    ...result,
    approval: {
      ...approvalForLive,
      confirmation_required: true,
      confirm_on_screen: true,
    },
  };
}

function juneSetConnected(connected) {
  const start = qs("juneStartBtn");
  const stop = qs("juneStopBtn");
  if (start) start.hidden = Boolean(connected);
  if (stop) stop.hidden = !connected;
}

function juneRenderConnectors(connectors) {
  const host = qs("juneConnectorList");
  if (!host) return;
  const labels = {
    calendar: "Calendar",
    email: "Email",
    weather: "Weather",
    ironlog: "Ironlog",
    borris: "Borris",
  };
  host.innerHTML = Object.entries(connectors || {}).map(([key, value]) => {
    const state = String(value?.state || "not_connected");
    const label = labels[key] || key;
    return `<div class="june-connector june-connector-${juneEscape(state)}"><span>${juneEscape(label)}</span><b>${juneEscape(state.replace(/_/g, " "))}</b></div>`;
  }).join("");
}

function juneRenderOutlookControl(outlook) {
  const control = qs("juneOutlookControl");
  const detail = qs("juneOutlookDetail");
  const connect = qs("juneConnectOutlookBtn");
  const disconnect = qs("juneDisconnectOutlookBtn");
  if (!control || !detail || !connect || !disconnect) return;
  const state = String(outlook?.state || "not_configured");
  const account = String(outlook?.account || "").trim();
  control.hidden = false;
  detail.textContent = account ? `${outlook?.detail || "Outlook connected."} (${account})` : String(outlook?.detail || "Outlook is not configured.");
  connect.hidden = !["not_connected", "needs_reconnect"].includes(state);
  connect.textContent = state === "needs_reconnect" ? "Reconnect Outlook" : "Connect Outlook";
  connect.disabled = state === "not_configured";
  disconnect.hidden = state !== "connected";
}

function juneCalendarSetFormVisible(visible) {
  const form = qs("juneCalendarForm");
  if (form) form.hidden = !visible;
}

function juneRenderCalendarEvents(events, { error = "" } = {}) {
  window.JuneCalendar?.setSource("ics", events, error);
}

function juneRenderIcsCalendar(calendar) {
  const card = qs("juneCalendarCard");
  const title = qs("juneCalendarTitle");
  const state = qs("juneCalendarState");
  const edit = qs("juneCalendarEditBtn");
  const refresh = qs("juneCalendarRefreshBtn");
  const remove = qs("juneCalendarRemoveBtn");
  if (!card || !title || !state || !edit || !refresh || !remove) return;
  juneIcsCalendar = calendar || null;
  const connection = String(calendar?.state || "not_connected");
  const name = String(calendar?.name || "My calendar").trim() || "My calendar";
  title.textContent = name;
  state.textContent = String(calendar?.last_error || calendar?.detail || "Add a private ICS link to see your upcoming meetings in Ironlog.");
  edit.title = connection === "connected" ? "Edit calendar link" : "Add your Outlook / private calendar link";
  remove.hidden = connection !== "connected";
  window.JuneCalendar?.setIcs({ connected: connection === "connected", name });
}

async function loadJuneStatus({ quiet = false } = {}) {
  if (!juneIsAdmin()) return;
  try {
    const data = await fetchJson(`${API}/api/june/status`);
    juneAvatarConfig = data?.avatar || null;
    const av = juneAvatarConfig;
    if (av && !av.configured) juneAvatarNote(`Animated avatar off: the server is missing ${(av.missing || []).join(", ")}.`);
    else if (av?.last_error) juneAvatarNote(`Animated avatar: LemonSlice refused the last session (HTTP ${av.last_error.status}${av.last_error.message ? `: ${av.last_error.message}` : ""}).`);
    else juneAvatarNote("");
    juneRenderConnectors(data?.connectors || {});
    juneRenderOutlookControl(data?.outlook || null);
    juneRenderIcsCalendar(data?.ics_calendar || null);
    if (!data?.live_ready) {
      juneSetState("June Live needs the server OpenAI configuration.", "warning");
      const start = qs("juneStartBtn");
      if (start) start.disabled = true;
      return;
    }
    const start = qs("juneStartBtn");
    if (start) start.disabled = false;
    const lastAttempt = data?.last_live_attempt;
    if (!junePeerConnection && lastAttempt?.state === "failed") {
      const detail = String(lastAttempt?.message || "June's last live session could not start.").trim();
      juneSetState(`June needs attention: ${detail}`, "warning");
      return;
    }
    if (!junePeerConnection && !quiet) juneSetState(`Ready — ${data?.voice || "gleam"} voice.`, "ready");
  } catch (error) {
    juneSetState("June is unavailable right now.", "warning");
  }
}

async function juneLoadIcsCalendar() {
  if (String(juneIcsCalendar?.state || "") !== "connected") return;
  const refresh = qs("juneCalendarRefreshBtn");
  if (refresh) refresh.disabled = true;
  try {
    const range = window.JuneCalendar?.range();
    const query = range ? `?from=${range.from}&to=${range.to}` : "";
    const data = await juneApiJson(`${API}/api/june/calendar/ics/events${query}`, { headers: authHeaders() });
    const calendar = data?.calendar || {};
    juneRenderIcsCalendar(calendar);
    juneRenderCalendarEvents(calendar?.events, { error: calendar?.error || "" });
  } catch (error) {
    juneRenderCalendarEvents([], { error: `Calendar could not be refreshed: ${error?.message || String(error)}` });
  } finally {
    if (refresh) refresh.disabled = false;
  }
}

async function juneSaveIcsCalendar(event) {
  event?.preventDefault?.();
  const url = String(qs("juneCalendarUrl")?.value || "").trim();
  const name = String(qs("juneCalendarName")?.value || "").trim();
  if (!url) return;
  const form = qs("juneCalendarForm");
  const submit = form?.querySelector('button[type="submit"]');
  if (submit) submit.disabled = true;
  try {
    const data = await juneApiJson(`${API}/api/june/calendar/ics`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ url, name }),
    });
    juneRenderIcsCalendar(data?.calendar || null);
    if (qs("juneCalendarUrl")) qs("juneCalendarUrl").value = "";
    juneCalendarSetFormVisible(false);
    await juneLoadIcsCalendar();
  } catch (error) {
    juneRenderCalendarEvents([], { error: `Calendar link was not saved: ${error?.message || String(error)}` });
  } finally {
    if (submit) submit.disabled = false;
  }
}

async function juneRemoveIcsCalendar() {
  const remove = qs("juneCalendarRemoveBtn");
  if (remove) remove.disabled = true;
  try {
    const data = await juneApiJson(`${API}/api/june/calendar/ics/remove`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: "{}",
    });
    juneRenderIcsCalendar(data?.calendar || null);
    juneCalendarSetFormVisible(false);
  } catch (error) {
    juneRenderCalendarEvents([], { error: `Calendar link could not be removed: ${error?.message || String(error)}` });
  } finally {
    if (remove) remove.disabled = false;
  }
}

async function juneConnectOutlook() {
  const button = qs("juneConnectOutlookBtn");
  if (button) button.disabled = true;
  juneSetState("June is opening Microsoft sign-in…", "working");
  try {
    const data = await juneApiJson(`${API}/api/june/outlook/connect`, { headers: authHeaders() });
    if (!data?.authorize_url) throw new Error("Ironlog did not receive a Microsoft authorisation link.");
    window.location.assign(data.authorize_url);
  } catch (error) {
    juneSetState(`Outlook could not start: ${error?.message || String(error)}`, "warning");
    if (button) button.disabled = false;
  }
}

async function juneDisconnectOutlook() {
  const button = qs("juneDisconnectOutlookBtn");
  if (button) button.disabled = true;
  juneSetState("Disconnecting Outlook…", "working");
  try {
    await juneApiJson(`${API}/api/june/outlook/disconnect`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: "{}",
    });
    await loadJuneStatus({ quiet: true });
    juneSetState("Outlook disconnected. June is back to keeping her nose out of your inbox.", "ready");
  } catch (error) {
    juneSetState(`Outlook could not be disconnected: ${error?.message || String(error)}`, "warning");
  } finally {
    if (button) button.disabled = false;
  }
}

function juneShowOutlookCallbackNotice() {
  const url = new URL(window.location.href);
  const state = String(url.searchParams.get("june_outlook") || "").trim();
  if (!state) return;
  const message = String(url.searchParams.get("june_outlook_message") || "Outlook connection updated.").trim();
  juneSetState(message, state === "connected" ? "ready" : "warning");
  url.searchParams.delete("june_outlook");
  url.searchParams.delete("june_outlook_message");
  window.history.replaceState({}, "", url.toString());
}

function syncJuneVisibility() {
  const card = qs("juneAssistantCard");
  const calendarCard = qs("juneCalendarCard");
  const row = qs("juneRow");
  if (!card) return;
  const visible = juneIsAdmin();
  card.hidden = !visible;
  if (row) row.hidden = !visible;
  // The Outlook-style calendar sits next to June: private calendar link + Ironlog schedule.
  if (calendarCard) calendarCard.hidden = !visible;
  if (!visible && junePeerConnection) juneStopLive({ silent: true });
  if (visible) {
    loadJuneStatus({ quiet: true })
      .catch(() => {})
      .then(() => window.JuneCalendar?.start(() => {
        juneLoadInternalCalendar({ quiet: true }).catch(() => {});
        juneLoadIcsCalendar().catch(() => {});
      }));
  }
}

function juneSendEvent(payload) {
  if (!juneDataChannel || juneDataChannel.readyState !== "open") return false;
  try {
    juneDataChannel.send(JSON.stringify(payload));
    return true;
  } catch {
    return false;
  }
}

function juneClientEventId(prefix) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

async function juneWaitForIceGathering(peerConnection, timeoutMs = 4_000) {
  if (peerConnection?.iceGatheringState === "complete") return;
  await new Promise((resolve) => {
    let timeoutId;
    const done = () => {
      clearTimeout(timeoutId);
      peerConnection?.removeEventListener("icegatheringstatechange", onStateChange);
      resolve();
    };
    const onStateChange = () => {
      if (peerConnection?.iceGatheringState === "complete") done();
    };
    timeoutId = window.setTimeout(done, timeoutMs);
    peerConnection?.addEventListener("icegatheringstatechange", onStateChange);
  });
}

async function juneApiJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  let data = {};
  try { data = JSON.parse(text); } catch {}
  if (!response.ok) {
    if (response.status === 401 && LOGIN_GATE_ENABLED) promptSignInAgain();
    const fallback = `June's voice service returned HTTP ${response.status}. Please retry shortly.`;
    throw Object.assign(new Error(String(data?.error || fallback)), { status: response.status });
  }
  return data;
}

function juneDelay(milliseconds) {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
}

async function juneCreateLiveSession(sdp) {
  const started = await juneApiJson(`${API}/api/june/live/session`, {
    method: "POST",
    headers: authHeaders({ "Content-Type": "application/json" }),
    body: JSON.stringify({ sdp }),
  });
  if (!started?.ticket || started?.state !== "pending") return started;
  const ticket = encodeURIComponent(String(started.ticket));
  const retryAfter = Math.max(300, Math.min(2_000, Number(started.retry_after_ms || 800)));
  for (let attempt = 0; attempt < 70; attempt += 1) {
    juneSetState("June is preparing her secure voice session…", "working");
    await juneDelay(retryAfter);
    const update = await juneApiJson(`${API}/api/june/live/session/${ticket}`, { headers: authHeaders() });
    if (update?.state === "ready" && update?.sdp) return update;
  }
  throw new Error("June's voice session is taking too long to initialise. Please retry.");
}

async function juneRunTool(item) {
  const callId = String(item?.call_id || "").trim();
  const name = String(item?.name || "").trim();
  if (!callId || !name || juneProcessedCalls.has(callId)) return;
  juneProcessedCalls.add(callId);
  let args = {};
  try { args = JSON.parse(String(item?.arguments || "{}")); } catch {}
  juneMemFlushUser();
  juneMemAdd("tool", `${name} ${JSON.stringify(args).slice(0, 200)}`);
  juneSetState("June is checking Ironlog…", "working");
  let output;
  try {
    const data = await fetchJson(`${API}/api/june/gateway/execute`, {
      method: "POST",
      body: JSON.stringify({ name, arguments: args }),
    });
    output = data?.result || { error: "June received no result from the gateway." };
    if (name === "june_draft_maintenance_schedule") {
      juneRenderMaintenanceScheduleDownload(output);
      // The opaque report id is only for the authenticated browser download.
      // June needs to know that the file is ready, never the identifier itself.
      output = {
        ...output,
        download: output?.download ? {
          available: true,
          label: output.download.label,
          expires_in_minutes: output.download.expires_in_minutes,
        } : null,
      };
    }
    if (name === "june_prepare_internal_calendar_change") {
      juneRenderCalendarApproval(output);
      // The opaque approval token belongs solely to the authenticated browser.
      // Keep it out of the live model's context so it cannot bypass the
      // explicit, visible confirmation step.
      output = juneCalendarToolOutputForLive(output);
    }
  } catch (error) {
    output = { error: `June could not complete ${name}: ${error?.message || String(error)}` };
  }
  const submitted = juneSendEvent({
    type: "response.item.create",
    event_id: juneClientEventId("june-tool"),
    item: {
      type: "function_call_output",
      call_id: callId,
      output: JSON.stringify(output),
    },
  });
  if (submitted) {
    juneSendEvent({ type: "response.create", event_id: juneClientEventId("june-continue") });
    juneSetState("June is preparing her answer…", "working");
  } else {
    juneSetState("June lost the live connection before the tool result returned.", "warning");
  }
}

function juneNestedLiveEvent(envelope) {
  const event = envelope?.type === "response.event" ? envelope.event : envelope;
  if (!event || typeof event !== "object") return;
  const type = String(event.type || "");
  if (type === "response.output_text.delta" || type === "response.output_audio_transcript.delta") {
    juneAppendTranscript("June", event.delta, "assistant");
    if (!juneMem.sawSessionOutput) {
      juneMemFlushUser();
      juneMem.june += String(event.delta || "");
    }
  }
  if (type === "response.created" || type === "response.output_item.added") {
    juneSetAvatarState("thinking");
  }
  if (type === "response.output_audio.delta") juneSetAvatarState("speaking");
  if (type === "response.done") {
    juneLemonSliceCommitResponse();
    juneMemFlushJune();
    return;
  }
  if (type === "input_audio_buffer.speech_started") {
    juneMemFlushJune();
    juneLemonSliceInterrupt();
    juneSetState("June is listening.", "live");
    return;
  }
  if (type === "input_audio_buffer.speech_stopped") {
    juneSetState("June is considering that…", "working");
    return;
  }
  if (type === "response.output_item.done" && String(event?.item?.type || "") === "function_call") {
    juneRunTool(event.item).catch(() => {});
  }
  if (type === "error") juneSetState(String(event?.error?.message || "June encountered a live-session error."), "warning");
}

function juneHandleLiveMessage(message) {
  let event;
  try { event = JSON.parse(message?.data || "{}"); } catch { return; }
  const type = String(event?.type || "");
  if (type === "session.started") {
    juneSetState("June is ready — start talking.", "ready");
    juneAppendTranscript("June", "June online. What are we taking on first?", "assistant");
    return;
  }
  if (type === "input_audio_buffer.speech_started") {
    juneMemFlushJune();
    juneLemonSliceInterrupt();
    juneSetState("June is listening.", "live");
    return;
  }
  if (type === "input_audio_buffer.speech_stopped") {
    juneSetState("June is considering that…", "working");
    return;
  }
  if (type === "session.input_transcript.delta") {
    juneMemFlushJune();
    juneMem.user += String(event.delta || "");
    juneAppendTranscript("You", event.delta, "user");
    juneSetAvatarState("listening");
    return;
  }
  if (type === "session.output_transcript.delta") {
    juneMem.sawSessionOutput = true;
    juneMemFlushUser();
    juneMem.june += String(event.delta || "");
    juneAppendTranscript("June", event.delta, "assistant");
    return;
  }
  if (type === "session.closed") {
    juneStopLive({ silent: true });
    juneSetState("June has ended the live conversation.", "neutral");
    return;
  }
  if (type === "error") {
    juneSetState(String(event?.error?.message || "June encountered a live-session error."), "warning");
    return;
  }
  juneNestedLiveEvent(event);
}

async function juneStartLive() {
  if (!juneIsAdmin()) return;
  if (!navigator.mediaDevices?.getUserMedia || !window.RTCPeerConnection) {
    juneSetState("This browser does not support the microphone connection June needs.", "warning");
    return;
  }
  const start = qs("juneStartBtn");
  const liveGeneration = ++juneLiveGeneration;
  if (start) start.disabled = true;
  juneClearTranscript();
  juneSetState("Requesting microphone access…", "working");
  try {
    juneLocalStream = await navigator.mediaDevices.getUserMedia({ audio: true });
    junePeerConnection = new RTCPeerConnection();
    juneAudio = new Audio();
    juneAudio.autoplay = true;
    junePeerConnection.ontrack = (event) => {
      const remoteStream = event.streams?.[0] || null;
      juneAudio.srcObject = remoteStream;
      juneRemoteAudioStream = remoteStream;
      if (remoteStream) {
        juneStartAvatarAudio(remoteStream);
        juneLemonSliceStartAudio(remoteStream);
      }
      juneAudio.play().catch(() => {});
    };
    juneAudio.addEventListener("playing", () => {
      juneSetAvatarState("speaking");
      if (juneAudio?.srcObject instanceof MediaStream) juneStartAvatarAudio(juneAudio.srcObject);
    });
    const returnToIdle = () => {
      if (junePeerConnection) juneSetAvatarState("idle");
    };
    juneAudio.addEventListener("pause", returnToIdle);
    juneAudio.addEventListener("ended", returnToIdle);
    junePeerConnection.onconnectionstatechange = () => {
      const state = String(junePeerConnection?.connectionState || "");
      if (state === "failed" || state === "disconnected") juneSetState("June’s voice connection was interrupted.", "warning");
    };
    juneLocalStream.getTracks().forEach((track) => junePeerConnection.addTrack(track, juneLocalStream));
    juneDataChannel = junePeerConnection.createDataChannel("oai-events");
    juneDataChannel.addEventListener("message", juneHandleLiveMessage);
    juneDataChannel.addEventListener("open", () => juneSetState("June is connecting…", "working"));
    juneDataChannel.addEventListener("close", () => {
      if (junePeerConnection) juneSetState("June’s live channel closed.", "neutral");
    });
    const offer = await junePeerConnection.createOffer();
    await junePeerConnection.setLocalDescription(offer);
    // OpenAI's WebRTC flow expects the SDP after candidate gathering has had a
    // chance to complete. Using localDescription also includes those candidates.
    await juneWaitForIceGathering(junePeerConnection);
    // Preserve the exact browser-produced SDP. The Live API accepts the SDP
    // offer verbatim; trimming its final CRLF can leave strict parsers at EOF.
    const sdp = junePeerConnection.localDescription?.sdp || offer.sdp || "";
    if (typeof sdp !== "string" || !sdp.startsWith("v=0") || !sdp.includes("m=audio")) {
      throw new Error("June could not create a valid WebRTC offer in this browser. Please retry after refreshing Ironlog.");
    }
    const response = await juneCreateLiveSession(sdp);
    await junePeerConnection.setRemoteDescription({ type: "answer", sdp: response.sdp });
    // A visual failure (or a slow visual provider) is non-fatal: never make
    // Jaco wait for June's already-working voice, intelligence, and tools.
    void juneStartLemonSliceAvatar(liveGeneration);
    juneSetConnected(true);
    juneSetState("June is starting…", "working");
  } catch (error) {
    juneStopLive({ silent: true });
    juneSetState(`June could not start: ${error?.message || String(error)}`, "warning");
  } finally {
    if (start && !junePeerConnection) start.disabled = false;
  }
}

function juneStopLive({ silent = false } = {}) {
  juneLiveGeneration += 1;
  juneMemFlushUser();
  juneMemFlushJune();
  juneMemSend({ final: true });
  juneMem.sawSessionOutput = false;
  juneSendEvent({ type: "session.close", event_id: juneClientEventId("june-close") });
  try { juneDataChannel?.close(); } catch {}
  try { junePeerConnection?.close(); } catch {}
  try { juneLocalStream?.getTracks()?.forEach((track) => track.stop()); } catch {}
  juneStopLemonSliceAvatar();
  if (juneAudio) {
    try { juneAudio.pause(); } catch {}
    juneAudio.srcObject = null;
  }
  juneStopAvatarAudio();
  juneSetAvatarState("idle");
  junePeerConnection = null;
  juneDataChannel = null;
  juneLocalStream = null;
  juneAudio = null;
  juneRemoteAudioStream = null;
  juneProcessedCalls = new Set();
  juneSetConnected(false);
  if (!silent) juneSetState("June is ready when you are.", "ready");
}

function initJune() {
  syncJuneVisibility();
  juneShowOutlookCallbackNotice();
}

function wireJuneControls() {
  qs("juneStartBtn")?.addEventListener("click", () => juneStartLive());
  qs("juneStopBtn")?.addEventListener("click", () => juneStopLive());
  qs("juneForgetBtn")?.addEventListener("click", () => juneForgetMemory());
  qs("juneRefreshBtn")?.addEventListener("click", () => loadJuneStatus());
  qs("juneConnectOutlookBtn")?.addEventListener("click", () => juneConnectOutlook());
  qs("juneDisconnectOutlookBtn")?.addEventListener("click", () => juneDisconnectOutlook());
  qs("juneCalendarEditBtn")?.addEventListener("click", () => {
    if (!qs("juneCalendarForm")?.hidden) return juneCalendarSetFormVisible(false);
    const name = qs("juneCalendarName");
    if (name) name.value = String(juneIcsCalendar?.name || "My calendar");
    juneCalendarSetFormVisible(true);
    qs("juneCalendarUrl")?.focus();
  });
  qs("juneCalendarCancelBtn")?.addEventListener("click", () => juneCalendarSetFormVisible(false));
  qs("juneCalendarForm")?.addEventListener("submit", juneSaveIcsCalendar);
  qs("juneCalendarRefreshBtn")?.addEventListener("click", () => {
    juneLoadIcsCalendar();
    juneLoadInternalCalendar();
  });
  qs("juneCalendarRemoveBtn")?.addEventListener("click", () => juneRemoveIcsCalendar());
  qs("juneInternalCalendarRefreshBtn")?.addEventListener("click", () => juneLoadInternalCalendar());
}


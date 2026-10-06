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

function juneRenderCalendarEvents(events, { error = "", emptyMessage = "No upcoming events in the next 21 days." } = {}) {
  const host = qs("juneCalendarEvents");
  if (!host) return;
  if (error) {
    host.innerHTML = `<div class="june-calendar-empty june-calendar-error">${juneEscape(error)}</div>`;
    return;
  }
  const list = Array.isArray(events) ? events : [];
  if (!list.length) {
    host.innerHTML = `<div class="june-calendar-empty">${juneEscape(emptyMessage)}</div>`;
    return;
  }
  host.innerHTML = list.map((event) => {
    const when = event?.all_day
      ? `${juneEscape(event?.start_date || "")} · All day`
      : `${juneEscape(event?.start_date || "")} · ${juneEscape(event?.start_time || "")}`;
    const location = String(event?.location || "").trim();
    return `<article class="june-calendar-event"><time>${when}</time><strong>${juneEscape(event?.summary || "(No title)")}</strong>${location ? `<span>${juneEscape(location)}</span>` : ""}</article>`;
  }).join("");
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
  edit.textContent = connection === "connected" ? "Edit calendar link" : "Add calendar link";
  refresh.hidden = connection !== "connected";
  remove.hidden = connection !== "connected";
  if (connection !== "connected") juneRenderCalendarEvents([], { emptyMessage: "Your upcoming calendar events will appear here." });
}

async function loadJuneStatus({ quiet = false } = {}) {
  if (!juneIsAdmin()) return;
  try {
    const data = await fetchJson(`${API}/api/june/status`);
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
  const refresh = qs("juneCalendarRefreshBtn");
  const state = qs("juneCalendarState");
  if (refresh) refresh.disabled = true;
  if (state) state.textContent = "Refreshing your private calendar…";
  try {
    const data = await juneApiJson(`${API}/api/june/calendar/ics/events`, { headers: authHeaders() });
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
  if (!card) return;
  const visible = juneIsAdmin();
  card.hidden = !visible;
  if (calendarCard) calendarCard.hidden = !visible;
  if (!visible && junePeerConnection) juneStopLive({ silent: true });
  if (visible) loadJuneStatus({ quiet: true }).catch(() => {});
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
  juneSetState("June is checking Ironlog…", "working");
  let output;
  try {
    const data = await fetchJson(`${API}/api/june/gateway/execute`, {
      method: "POST",
      body: JSON.stringify({ name, arguments: args }),
    });
    output = data?.result || { error: "June received no result from the gateway." };
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
    juneSetState("June is listening.", "live");
    juneAppendTranscript("June", "June online. What are we taking on first?", "assistant");
    return;
  }
  if (type === "session.input_transcript.delta") {
    juneAppendTranscript("You", event.delta, "user");
    return;
  }
  if (type === "session.output_transcript.delta") {
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
  if (start) start.disabled = true;
  juneClearTranscript();
  juneSetState("Requesting microphone access…", "working");
  try {
    juneLocalStream = await navigator.mediaDevices.getUserMedia({ audio: true });
    junePeerConnection = new RTCPeerConnection();
    juneAudio = new Audio();
    juneAudio.autoplay = true;
    junePeerConnection.ontrack = (event) => {
      juneAudio.srcObject = event.streams?.[0] || null;
      juneAudio.play().catch(() => {});
    };
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
  juneSendEvent({ type: "session.close", event_id: juneClientEventId("june-close") });
  try { juneDataChannel?.close(); } catch {}
  try { junePeerConnection?.close(); } catch {}
  try { juneLocalStream?.getTracks()?.forEach((track) => track.stop()); } catch {}
  if (juneAudio) {
    try { juneAudio.pause(); } catch {}
    juneAudio.srcObject = null;
  }
  junePeerConnection = null;
  juneDataChannel = null;
  juneLocalStream = null;
  juneAudio = null;
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
  qs("juneRefreshBtn")?.addEventListener("click", () => loadJuneStatus());
  qs("juneConnectOutlookBtn")?.addEventListener("click", () => juneConnectOutlook());
  qs("juneDisconnectOutlookBtn")?.addEventListener("click", () => juneDisconnectOutlook());
  qs("juneCalendarEditBtn")?.addEventListener("click", () => {
    const name = qs("juneCalendarName");
    if (name) name.value = String(juneIcsCalendar?.name || "My calendar");
    juneCalendarSetFormVisible(true);
    qs("juneCalendarUrl")?.focus();
  });
  qs("juneCalendarCancelBtn")?.addEventListener("click", () => juneCalendarSetFormVisible(false));
  qs("juneCalendarForm")?.addEventListener("submit", juneSaveIcsCalendar);
  qs("juneCalendarRefreshBtn")?.addEventListener("click", () => juneLoadIcsCalendar());
  qs("juneCalendarRemoveBtn")?.addEventListener("click", () => juneRemoveIcsCalendar());
}


// IRONLOG/web/maintenance/sync.js — InspectPro sync administration.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

function syncSetMsg(text, isError = false) {
  const el = document.getElementById("syncMsg");
  if (!el) return;
  el.className = isError ? "message-error" : "muted";
  el.textContent = text || "";
}

function syncSetOutput(obj) {
  const el = document.getElementById("syncOutput");
  if (!el) return;
  try {
    el.textContent = JSON.stringify(obj ?? {}, null, 2);
  } catch {
    el.textContent = String(obj ?? "");
  }
}

function syncHeaders() {
  return authHeaders({ "Content-Type": "application/json" });
}


function getSyncForm() {
  const peer = String(document.getElementById("syncPeer")?.value || "").trim();
  const since = Number(document.getElementById("syncSinceId")?.value || 0);
  const limit = Number(document.getElementById("syncLimit")?.value || 200);
  return {
    peer: peer || "local-maint-ui",
    since_id: Number.isFinite(since) && since >= 0 ? Math.trunc(since) : 0,
    limit: Number.isFinite(limit) ? Math.max(1, Math.min(5000, Math.trunc(limit))) : 200,
  };
}

async function syncLoadStats() {
  syncSetMsg("Loading sync stats...");
  try {
    const res = await fetch(`${API}/sync/stats`, { headers: syncHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Stats load failed");
    syncSetOutput(data);
    syncSetMsg(`Stats loaded. Unsynced: ${Number(data.unsynced || 0)}.`);
  } catch (e) {
    syncSetMsg(`Sync stats error: ${e.message || e}`, true);
  }
}

async function syncLoadState() {
  const f = getSyncForm();
  syncSetMsg("Loading sync state...");
  try {
    const q = new URLSearchParams();
    q.set("schema_version", "1");
    if (f.peer) q.set("peer", f.peer);
    const res = await fetch(`${API}/sync/state?${q.toString()}`, { headers: syncHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Sync state load failed");
    syncSetOutput(data);
    syncSetMsg(`State loaded${data.peer ? ` for ${data.peer}` : ""}.`);
  } catch (e) {
    syncSetMsg(`Sync state error: ${e.message || e}`, true);
  }
}

async function syncLoadOutbox() {
  const { limit } = getSyncForm();
  syncSetMsg("Loading outbox...");
  try {
    const res = await fetch(`${API}/sync/outbox?limit=${limit}`, { headers: syncHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Outbox load failed");
    syncSetOutput(data);
    syncSetMsg(`Outbox loaded (${Number(data.count || 0)} rows).`);
  } catch (e) {
    syncSetMsg(`Sync outbox error: ${e.message || e}`, true);
  }
}

async function syncPullEvents() {
  const f = getSyncForm();
  syncSetMsg("Pulling sync events...");
  try {
    const q = new URLSearchParams();
    q.set("peer", f.peer);
    q.set("since_id", String(f.since_id));
    q.set("limit", String(f.limit));
    const res = await fetch(`${API}/sync/pull?${q.toString()}`, { headers: syncHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Sync pull failed");
    lastSyncPull = { last_id: Number(data.last_id || 0), events: Array.isArray(data.events) ? data.events : [] };
    const sinceEl = document.getElementById("syncSinceId");
    if (sinceEl) sinceEl.value = String(lastSyncPull.last_id || f.since_id);
    syncSetOutput(data);
    syncSetMsg(`Pulled ${lastSyncPull.events.length} event(s). Last id: ${lastSyncPull.last_id}.`);
  } catch (e) {
    syncSetMsg(`Sync pull error: ${e.message || e}`, true);
  }
}

async function syncApplyLastPull() {
  const f = getSyncForm();
  if (!lastSyncPull.events.length) {
    syncSetMsg("No pulled events to apply yet. Click Pull first.", true);
    return;
  }
  syncSetMsg(`Applying ${lastSyncPull.events.length} event(s)...`);
  try {
    const res = await fetch(`${API}/sync/apply`, {
      method: "POST",
      headers: syncHeaders(),
      body: JSON.stringify({ peer: f.peer, events: lastSyncPull.events }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Sync apply failed");
    syncSetOutput(data);
    syncSetMsg(`Apply complete. Applied: ${Number(data.applied || 0)}, skipped: ${Number(data.skipped || 0)}.`);
  } catch (e) {
    syncSetMsg(`Sync apply error: ${e.message || e}`, true);
  }
}

async function syncApplyLastPullDryRun() {
  const f = getSyncForm();
  if (!lastSyncPull.events.length) {
    syncSetMsg("No pulled events to dry-run. Click Pull first.", true);
    return;
  }
  syncSetMsg(`Dry-run applying ${lastSyncPull.events.length} event(s)...`);
  try {
    const res = await fetch(`${API}/sync/apply`, {
      method: "POST",
      headers: syncHeaders(),
      body: JSON.stringify({ schema_version: 1, dry_run: 1, peer: f.peer, events: lastSyncPull.events }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Sync apply dry-run failed");
    syncSetOutput(data);
    syncSetMsg(`Dry-run done. Would apply: ${Number(data.would_apply || 0)}, skipped: ${Number(data.skipped || 0)}.`);
  } catch (e) {
    syncSetMsg(`Sync apply dry-run error: ${e.message || e}`, true);
  }
}

function syncExportLastPull() {
  const f = getSyncForm();
  if (!lastSyncPull.events.length) {
    syncSetMsg("No pulled events to export yet. Click Pull first.", true);
    return;
  }
  try {
    const payload = {
      schema_version: 1,
      exported_at: new Date().toISOString(),
      peer: f.peer,
      last_id: Number(lastSyncPull.last_id || 0),
      count: Number(lastSyncPull.events.length || 0),
      events: lastSyncPull.events,
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const stamp = new Date().toISOString().replace(/[:.]/g, "-");
    const a = document.createElement("a");
    a.href = url;
    a.download = `ironlog_sync_last_pull_${f.peer || "peer"}_${stamp}.json`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    syncSetMsg(`Exported ${lastSyncPull.events.length} event(s) to JSON.`);
  } catch (e) {
    syncSetMsg(`Export failed: ${e.message || e}`, true);
  }
}

async function syncAckLastPull() {
  if (!lastSyncPull.events.length) {
    syncSetMsg("No pulled events to acknowledge yet. Click Pull first.", true);
    return;
  }
  const ids = lastSyncPull.events
    .map((e) => Number(e?.id))
    .filter((n) => Number.isInteger(n) && n > 0);
  if (!ids.length) {
    syncSetMsg("Last pull has no event IDs to acknowledge.", true);
    return;
  }
  syncSetMsg(`Acknowledging ${ids.length} outbox event(s)...`);
  try {
    const res = await fetch(`${API}/sync/outbox/ack`, {
      method: "POST",
      headers: syncHeaders(),
      body: JSON.stringify({ ids }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Outbox ack failed");
    syncSetOutput(data);
    syncSetMsg(`Acknowledged ${Number(data.acknowledged || 0)} outbox event(s).`);
  } catch (e) {
    syncSetMsg(`Outbox ack error: ${e.message || e}`, true);
  }
}

async function syncCheckpointLastPull() {
  const f = getSyncForm();
  const lastId = Number(lastSyncPull.last_id || f.since_id || 0);
  syncSetMsg(`Saving checkpoint at ${lastId}...`);
  try {
    const res = await fetch(`${API}/sync/checkpoint`, {
      method: "POST",
      headers: syncHeaders(),
      body: JSON.stringify({ peer: f.peer, last_outbox_id: lastId }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Checkpoint save failed");
    syncSetOutput(data);
    syncSetMsg(`Checkpoint saved for ${f.peer} at ${lastId}.`);
  } catch (e) {
    syncSetMsg(`Checkpoint error: ${e.message || e}`, true);
  }
}

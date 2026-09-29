// IRONLOG/web/maintenance/histogram.js — Maintenance histogram events.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

async function loadHistogramEvents() {
  const body = document.getElementById("histEventBody");
  const msg = document.getElementById("histMsg");
  if (!body) return;
  body.innerHTML = `<tr><td colspan="10" class="muted">Loading...</td></tr>`;
  try {
    const q = new URLSearchParams();
    const start = String(document.getElementById("histFilterStart")?.value || "").trim();
    const end = String(document.getElementById("histFilterEnd")?.value || "").trim();
    const location = String(document.getElementById("histFilterLocation")?.value || "").trim();
    const part = String(document.getElementById("histFilterPart")?.value || "").trim();
    const approval = String(document.getElementById("histFilterApproval")?.value || "").trim();
    if (start) q.set("start", start);
    if (end) q.set("end", end);
    if (location) q.set("location", location);
    if (part) q.set("part", part);
    if (approval) q.set("approval", approval);
    const res = await fetch(`${API}/maintenance/histogram/events?${q.toString()}`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load histogram events");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    if (!rows.length) {
      body.innerHTML = `<tr><td colspan="10" class="muted">No events found for selected filters.</td></tr>`;
      if (msg) {
        msg.className = "muted";
        msg.textContent = "No events found.";
      }
      return;
    }
    body.innerHTML = rows.map((r) => `
      <tr>
        <td>${esc(r.event_date || "-")}</td>
        <td>${esc(r.asset_number || "-")}</td>
        <td>${esc(r.location || "-")}</td>
        <td>${esc(r.part_code || "-")}</td>
        <td>${esc(r.part_name || "-")}</td>
        <td>${esc(r.approval_status || "-")}</td>
        <td>${esc(r.approved_by || "-")}</td>
        <td>${esc(r.notes || "-")}</td>
        <td>${esc(r.created_by || "-")}</td>
        <td style="white-space:nowrap;">
          <button type="button" data-hist-edit="${Number(r.id || 0)}">Edit</button>
          <button type="button" data-hist-del="${Number(r.id || 0)}">Delete</button>
        </td>
      </tr>
    `).join("");
    if (msg) {
      msg.className = "muted";
      msg.textContent = `Loaded ${rows.length} event(s).`;
    }
  } catch (e) {
    body.innerHTML = `<tr><td colspan="10" class="message-error">${esc(e.message || String(e))}</td></tr>`;
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Load error: ${e.message || e}`;
    }
  }
}

async function saveHistogramEvent() {
  const msg = document.getElementById("histMsg");
  const event_date = String(document.getElementById("histEventDate")?.value || "").trim();
  const asset_number = String(document.getElementById("histAssetNumber")?.value || "").trim();
  const location = String(document.getElementById("histLocation")?.value || "").trim();
  const part_code = String(document.getElementById("histPartCode")?.value || "").trim();
  const part_name = String(document.getElementById("histPartName")?.value || "").trim();
  const approval_status = String(document.getElementById("histApprovalStatus")?.value || "").trim();
  const approved_by = String(document.getElementById("histApprovedBy")?.value || "").trim();
  const notes = String(document.getElementById("histNotes")?.value || "").trim();
  if (!event_date) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Event date is required.";
    }
    return;
  }
  try {
    const res = await fetch(`${API}/maintenance/histogram/events`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ event_date, asset_number, location, part_code, part_name, approval_status, approved_by, notes }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save event");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Histogram event saved.";
    }
    const clearIds = ["histAssetNumber", "histLocation", "histPartCode", "histPartName", "histApprovedBy", "histNotes"];
    clearIds.forEach((id) => {
      const el = document.getElementById(id);
      if (el) el.value = "";
    });
    const approvalEl = document.getElementById("histApprovalStatus");
    if (approvalEl) approvalEl.value = "";
    await loadHistogramEvents();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Save error: ${e.message || e}`;
    }
  }
}

function openHistogramPdf(download = false) {
  const q = new URLSearchParams();
  q.set("include_all", "1");
  q.set("site_code", getSessionSite());
  if (download) q.set("download", "1");
  return openProtectedPdf(`${API}/maintenance/histogram/events.pdf?${q.toString()}`, {
    download,
    filename: "IRONLOG_Histogram_Events.pdf",
  });
}

async function editHistogramEvent(id) {
  const n = Number(id || 0);
  if (!n) return;
  const rows = Array.from(document.querySelectorAll("#histEventBody tr"));
  const row = rows.find((tr) => Number(tr.querySelector("button[data-hist-edit]")?.getAttribute("data-hist-edit") || 0) === n);
  const tds = row ? row.querySelectorAll("td") : [];
  const currentDate = String(tds[0]?.textContent || "").trim();
  const currentAssetNumber = String(tds[1]?.textContent || "").trim();
  const currentLocation = String(tds[2]?.textContent || "").trim();
  const currentPartCode = String(tds[3]?.textContent || "").trim();
  const currentPartName = String(tds[4]?.textContent || "").trim();
  const currentApproval = String(tds[5]?.textContent || "").trim();
  const currentApprovedBy = String(tds[6]?.textContent || "").trim();
  const currentNotes = String(tds[7]?.textContent || "").trim();

  const event_date = String(window.prompt("Event date (YYYY-MM-DD):", currentDate) || "").trim();
  if (!event_date) return;
  const asset_number = String(window.prompt("Asset number:", currentAssetNumber) || "").trim();
  const location = String(window.prompt("Location:", currentLocation) || "").trim();
  const part_code = String(window.prompt("Part code:", currentPartCode) || "").trim();
  const part_name = String(window.prompt("Part name:", currentPartName) || "").trim();
  const approval_status = String(window.prompt("Approval status (Pending/Approved/Rejected):", currentApproval) || "").trim();
  const approved_by = String(window.prompt("Approved by:", currentApprovedBy) || "").trim();
  const notes = String(window.prompt("Notes:", currentNotes) || "").trim();

  const msg = document.getElementById("histMsg");
  try {
    const res = await fetch(`${API}/maintenance/histogram/events/${n}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ event_date, asset_number, location, part_code, part_name, approval_status, approved_by, notes }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to update event");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Histogram event updated.";
    }
    await loadHistogramEvents();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Update error: ${e.message || e}`;
    }
  }
}

async function deleteHistogramEvent(id) {
  const n = Number(id || 0);
  if (!n) return;
  if (!window.confirm("Delete this histogram event?")) return;
  const msg = document.getElementById("histMsg");
  try {
    const res = await fetch(`${API}/maintenance/histogram/events/${n}`, {
      method: "DELETE",
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to delete event");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = "Histogram event deleted.";
    }
    await loadHistogramEvents();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = `Delete error: ${e.message || e}`;
    }
  }
}

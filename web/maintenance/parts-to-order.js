// IRONLOG/web/maintenance/parts-to-order.js — Parts to order and RFQ.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

const PTO_MANAGER_ROLES = new Set(["admin", "supervisor", "stores", "plant_manager", "site_manager", "workshop_manager"]);

function canManagePartsOrderStatus() {
  return getSessionRoles().some((r) => PTO_MANAGER_ROLES.has(r));
}

function isArtisanOnlySession() {
  const roles = getSessionRoles();
  return roles.length === 1 && roles[0] === "artisan";
}

function ptoStatusLabel(status) {
  const s = String(status || "").toLowerCase();
  if (s === "requested") return "Pending";
  if (s === "ordered") return "Ordered";
  if (s === "received") return "Received";
  if (s === "cancelled") return "Cancelled";
  return s || "—";
}

function ptoUrgencyLabel(urgency) {
  const u = String(urgency || "normal").toLowerCase();
  if (u === "critical") return "Critical";
  if (u === "urgent") return "Urgent";
  return "Normal";
}

function ptoUrgencyClass(urgency) {
  const u = String(urgency || "normal").toLowerCase();
  if (u === "critical") return "message-error";
  if (u === "urgent") return "message-error";
  return "muted";
}

let ptoRowsCache = [];
const selectedPtoRequestIds = new Set();

function defaultPtoRfqReference() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const date = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}`;
  const site = getSessionSite().toUpperCase().replace(/[^A-Z0-9]/g, "") || "MAIN";
  return `RFQ-${site}-${date}`;
}

function initPtoRfqDefaults() {
  const ref = document.getElementById("ptoRfqReference");
  if (ref && !String(ref.value || "").trim()) ref.value = defaultPtoRfqReference();
  const contact = document.getElementById("ptoRfqContact");
  if (contact && !String(contact.value || "").trim()) contact.value = getSessionUser();
  const reqBy = document.getElementById("ptoRfqRequiredBy");
  if (reqBy && !reqBy.value) {
    const d = new Date();
    d.setDate(d.getDate() + 7);
    reqBy.value = d.toISOString().slice(0, 10);
  }
}

function updatePtoSelectionCount() {
  const el = document.getElementById("ptoSelectionCount");
  if (el) el.textContent = `${selectedPtoRequestIds.size} selected`;
  const hint = document.getElementById("ptoRfqSelectionHint");
  if (hint) {
    hint.textContent = selectedPtoRequestIds.size
      ? `${selectedPtoRequestIds.size} part line(s) selected for RFQ`
      : "Tick parts above, then complete supplier details.";
  }
  renderPtoRfqPreview();
}

function bindPtoRequestCheckboxes(root) {
  (root || document).querySelectorAll(".pto-request-select").forEach((el) => {
    if (el.dataset.boundPtoSelect === "1") return;
    el.dataset.boundPtoSelect = "1";
    el.addEventListener("change", () => {
      const id = Number(el.getAttribute("data-pto-id") || 0);
      if (!id) return;
      if (el.checked) selectedPtoRequestIds.add(id);
      else selectedPtoRequestIds.delete(id);
      updatePtoSelectionCount();
    });
  });
  updatePtoSelectionCount();
}

function clearPtoSelection() {
  selectedPtoRequestIds.clear();
  document.querySelectorAll(".pto-request-select").forEach((el) => { el.checked = false; });
  updatePtoSelectionCount();
}

function selectAllPtoRequests() {
  document.querySelectorAll("#ptoListBody .pto-request-select").forEach((el) => {
    const id = Number(el.getAttribute("data-pto-id") || 0);
    if (!id) return;
    el.checked = true;
    selectedPtoRequestIds.add(id);
  });
  updatePtoSelectionCount();
}

function getSelectedPtoRows() {
  const ids = selectedPtoRequestIds;
  return (ptoRowsCache || []).filter((r) => ids.has(Number(r.id || 0)));
}

function renderPtoRfqPreview() {
  const body = document.getElementById("ptoRfqPreviewBody");
  if (!body) return;
  const rows = getSelectedPtoRows();
  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="7" class="muted">No parts selected.</td></tr>`;
    return;
  }
  body.innerHTML = rows.map((r, idx) => {
    const asset = r.asset_code ? esc(r.asset_code) : "—";
    return `
      <tr>
        <td>${idx + 1}</td>
        <td>${r.part_code ? esc(r.part_code) : "—"}</td>
        <td>${esc(r.part_name || "")}</td>
        <td style="text-align:right;">${Number(r.qty || 0)}</td>
        <td>${asset}</td>
        <td>${r.work_order_id ? esc(String(r.work_order_id)) : "—"}</td>
        <td>${esc(ptoUrgencyLabel(r.urgency))}</td>
      </tr>
    `;
  }).join("");
}

function buildPtoRfqPdfQuery() {
  const ids = Array.from(selectedPtoRequestIds);
  if (!ids.length) throw new Error("Select at least one part request for the RFQ.");
  const supplier = String(document.getElementById("ptoRfqSupplier")?.value || "").trim();
  if (!supplier) throw new Error("Enter a supplier / company name.");
  const q = new URLSearchParams();
  q.set("ids", ids.join(","));
  q.set("reference", String(document.getElementById("ptoRfqReference")?.value || "").trim() || defaultPtoRfqReference());
  q.set("supplier", supplier);
  const contact = String(document.getElementById("ptoRfqContact")?.value || "").trim();
  const email = String(document.getElementById("ptoRfqEmail")?.value || "").trim();
  const phone = String(document.getElementById("ptoRfqPhone")?.value || "").trim();
  const required_by = String(document.getElementById("ptoRfqRequiredBy")?.value || "").trim();
  const notes = String(document.getElementById("ptoRfqNotes")?.value || "").trim();
  if (contact) q.set("contact", contact);
  if (email) q.set("email", email);
  if (phone) q.set("phone", phone);
  if (required_by) q.set("required_by", required_by);
  if (notes) q.set("notes", notes);
  q.set("requested_by", getSessionUser());
  q.set("_", String(Date.now()));
  return q;
}

async function openPtoRfqPdf(download = false) {
  const msg = document.getElementById("ptoRfqMsg");
  try {
    const q = buildPtoRfqPdfQuery();
    if (download) q.set("download", "1");
    const url = `${API}/maintenance/parts-requests/rfq.pdf?${q.toString()}`;
    if (msg) {
      msg.className = "muted";
      msg.textContent = download ? "Preparing download..." : "Opening RFQ PDF...";
    }
    const res = await fetch(url, { headers: authHeaders() });
    if (!res.ok) {
      const t = await res.text();
      throw new Error(t || `RFQ PDF failed (${res.status})`);
    }
    const blob = await res.blob();
    const blobUrl = URL.createObjectURL(blob);
    const ref = String(document.getElementById("ptoRfqReference")?.value || "").trim() || "rfq";
    if (download) {
      const a = document.createElement("a");
      a.href = blobUrl;
      a.download = `${ref.replace(/[^\w.-]+/g, "_")}.pdf`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(blobUrl), 3000);
    } else {
      window.open(blobUrl, "_blank");
      setTimeout(() => URL.revokeObjectURL(blobUrl), 15000);
    }
    if (msg) {
      msg.className = "message-success";
      msg.textContent = download ? "RFQ PDF downloaded." : "RFQ PDF opened.";
    }
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    } else {
      alert(e.message || String(e));
    }
  }
}

function clearPartsToOrderForm() {
  const asset = document.getElementById("ptoAsset");
  const code = document.getElementById("ptoPartCode");
  const name = document.getElementById("ptoPartName");
  const qty = document.getElementById("ptoQty");
  const urgency = document.getElementById("ptoUrgency");
  const wo = document.getElementById("ptoWorkOrderId");
  const notes = document.getElementById("ptoNotes");
  const msg = document.getElementById("ptoFormMsg");
  if (asset) asset.value = "";
  if (code) code.value = "";
  if (name) name.value = "";
  if (qty) qty.value = "1";
  if (urgency) urgency.value = "normal";
  if (wo) wo.value = "";
  if (notes) notes.value = "";
  if (msg) {
    msg.className = "muted";
    msg.textContent = "";
  }
}

function bindPtoPartCodeAutofill() {
  const input = document.getElementById("ptoPartCode");
  if (!input || input.dataset.ptoBound === "1") return;
  input.dataset.ptoBound = "1";
  input.addEventListener("change", () => {
    const hit = getWfPartByCode(input.value);
    if (!hit) return;
    const nameEl = document.getElementById("ptoPartName");
    if (nameEl && !String(nameEl.value || "").trim()) {
      nameEl.value = String(hit.part_name || "").trim();
    }
  });
}

function renderPartsToOrderActions(row) {
  const id = Number(row.id || 0);
  const status = String(row.status || "requested").toLowerCase();
  const requestedBy = String(row.requested_by || "").trim().toLowerCase();
  const me = getSessionUser().toLowerCase();
  const own = requestedBy && requestedBy === me;
  const manager = canManagePartsOrderStatus();
  const bits = [];

  if (manager && status === "requested") {
    bits.push(`<button type="button" class="btn-sm" data-pto-status="${id}" data-pto-next="ordered">Mark ordered</button>`);
  }
  if (manager && status === "ordered") {
    bits.push(`<button type="button" class="btn-sm" data-pto-status="${id}" data-pto-next="received">Mark received</button>`);
  }
  if ((manager || own) && (status === "requested" || (manager && status === "ordered"))) {
    bits.push(`<button type="button" class="btn-sm btn-secondary" data-pto-status="${id}" data-pto-next="cancelled">Cancel</button>`);
  }
  return bits.join(" ");
}

function renderPartsToOrderTable(rows) {
  const body = document.getElementById("ptoListBody");
  if (!body) return;
  ptoRowsCache = Array.isArray(rows) ? rows : [];
  const validIds = new Set(ptoRowsCache.map((r) => Number(r.id || 0)).filter((x) => x > 0));
  Array.from(selectedPtoRequestIds).forEach((id) => {
    if (!validIds.has(id)) selectedPtoRequestIds.delete(id);
  });
  if (!ptoRowsCache.length) {
    body.innerHTML = `<tr><td colspan="11" class="muted">No parts requests found.</td></tr>`;
    updatePtoSelectionCount();
    return;
  }
  body.innerHTML = ptoRowsCache.map((r) => {
    const id = Number(r.id || 0);
    const asset = r.asset_code
      ? `${esc(r.asset_code)}${r.asset_name ? ` — ${esc(r.asset_name)}` : ""}`
      : "—";
    const partBits = [
      r.part_code ? `<strong>${esc(r.part_code)}</strong>` : "",
      esc(r.part_name || ""),
    ].filter(Boolean).join("<br />");
    const noteBits = [r.notes, r.status_notes].map((x) => String(x || "").trim()).filter(Boolean);
    return `
      <tr data-pto-id="${id}">
        <td style="text-align:center;">
          ${id > 0 ? `<input type="checkbox" class="pto-request-select" data-pto-id="${id}" title="Select for RFQ" ${selectedPtoRequestIds.has(id) ? "checked" : ""} />` : ""}
        </td>
        <td>${esc(String(r.created_at || "").slice(0, 16).replace("T", " "))}</td>
        <td>${asset}</td>
        <td>${partBits || "—"}</td>
        <td style="text-align:right;">${Number(r.qty || 0)}</td>
        <td><span class="${ptoUrgencyClass(r.urgency)}">${esc(ptoUrgencyLabel(r.urgency))}</span></td>
        <td>${r.work_order_id ? esc(String(r.work_order_id)) : "—"}</td>
        <td>${esc(r.requested_by || "")}</td>
        <td>${esc(ptoStatusLabel(r.status))}${r.ordered_by ? `<br /><small class="muted">${esc(r.ordered_by)}</small>` : ""}</td>
        <td>${noteBits.length ? esc(noteBits.join(" · ")) : "—"}</td>
        <td style="white-space:nowrap;">${renderPartsToOrderActions(r)}</td>
      </tr>
    `;
  }).join("");
  bindPtoRequestCheckboxes(body);
}

async function loadPartsToOrderList() {
  const body = document.getElementById("ptoListBody");
  if (body) body.innerHTML = `<tr><td colspan="11" class="muted">Loading...</td></tr>`;
  const status = String(document.getElementById("ptoFilterStatus")?.value || "").trim();
  const mine = document.getElementById("ptoMineOnly")?.checked ? "1" : "";
  const q = new URLSearchParams();
  if (status) q.set("status", status);
  if (mine) q.set("mine", mine);
  try {
    const res = await fetch(`${API}/maintenance/parts-requests?${q.toString()}`, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load parts requests");
    let rows = Array.isArray(data.rows) ? data.rows : [];
    if (!status) {
      rows = rows.filter((r) => String(r.status || "").toLowerCase() !== "cancelled" && String(r.status || "").toLowerCase() !== "received");
    }
    renderPartsToOrderTable(rows);
  } catch (e) {
    if (body) body.innerHTML = `<tr><td colspan="11" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

async function submitPartsToOrderRequest() {
  const msg = document.getElementById("ptoFormMsg");
  const asset_id = Number(document.getElementById("ptoAsset")?.value || 0) || null;
  const part_code = String(document.getElementById("ptoPartCode")?.value || "").trim();
  const part_name = String(document.getElementById("ptoPartName")?.value || "").trim();
  const qty = Number(document.getElementById("ptoQty")?.value || 1);
  const urgency = String(document.getElementById("ptoUrgency")?.value || "normal").trim();
  const work_order_id = Number(document.getElementById("ptoWorkOrderId")?.value || 0) || null;
  const notes = String(document.getElementById("ptoNotes")?.value || "").trim();
  if (!msg) return;
  if (!part_code && !part_name) {
    msg.className = "message-error";
    msg.textContent = "Enter a part code or description.";
    return;
  }
  if (!Number.isFinite(qty) || qty <= 0) {
    msg.className = "message-error";
    msg.textContent = "Quantity must be greater than zero.";
    return;
  }
  msg.className = "muted";
  msg.textContent = "Submitting request...";
  try {
    const res = await fetch(`${API}/maintenance/parts-requests`, {
      method: "POST",
      headers: authHeaders({ "Content-Type": "application/json" }),
      body: JSON.stringify({ asset_id, part_code, part_name, qty, urgency, work_order_id, notes }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to submit request");
    msg.className = "message-success";
    msg.textContent = "Part request submitted.";
    clearPartsToOrderForm();
    await loadPartsToOrderList();
  } catch (e) {
    msg.className = "message-error";
    msg.textContent = e.message || String(e);
  }
}

async function updatePartsToOrderStatus(id, status) {
  const n = Number(id || 0);
  if (!n) return;
  let status_notes = "";
  if (status === "cancelled") {
    status_notes = String(window.prompt("Cancellation reason (optional):", "") || "").trim();
  } else if (status === "ordered") {
    status_notes = String(window.prompt("PO / supplier note (optional):", "") || "").trim();
  }
  const res = await fetch(`${API}/maintenance/parts-requests/${n}/status`, {
    method: "PATCH",
    headers: authHeaders({ "Content-Type": "application/json" }),
    body: JSON.stringify({ status, status_notes }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Failed to update status");
  await loadPartsToOrderList();
}

async function initPartsToOrderSection() {
  bindPtoPartCodeAutofill();
  initPtoRfqDefaults();
  const mine = document.getElementById("ptoMineOnly");
  if (mine && !mine.dataset.inited) {
    mine.dataset.inited = "1";
    mine.checked = isArtisanOnlySession();
  }
  await loadWeeklyForumParts();
  await loadPartsToOrderList();
}

// IRONLOG/web/maintenance/inspections.js — Manager, damage and artisan inspections.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

async function saveManagerInspection() {
  const asset_id = Number(document.getElementById("miAsset")?.value || 0);
  const inspection_date = String(document.getElementById("miDate")?.value || "").trim();
  const inspector_name = String(document.getElementById("miInspector")?.value || "").trim();
  const notes = String(document.getElementById("miNotes")?.value || "").trim();
  const msg = document.getElementById("miMsg");
  const mhRaw = String(document.getElementById("miMachineHours")?.value || "").trim();
  const machine_hours = mhRaw === "" ? null : Number(mhRaw);
  const inspection_type = getManagerInspectionType();
  if (machine_hours != null && !Number.isFinite(machine_hours)) {
    return alert("Machine hours must be a number.");
  }
  const checklist = collectManagerInspectionChecklist();
  const required_parts = collectManagerInspectionParts();
  const failedRows = getFailedChecklistRows(checklist);
  const requireEvidence = document.getElementById("miRequireEvidence")?.checked !== false;
  const evidenceFiles = Array.from(document.getElementById("miEvidencePhotos")?.files || []);
  const defect_component = String(document.getElementById("miDefectComponent")?.value || "").trim();
  const defect_severity = String(document.getElementById("miDefectSeverity")?.value || "").trim().toLowerCase();
  const defect_risk = String(document.getElementById("miDefectRisk")?.value || "").trim();
  const recommended_action = String(document.getElementById("miRecommendedAction")?.value || "").trim();
  const create_work_order = document.getElementById("miCreateWoAlways")?.checked === true;
  const create_work_order_on_issues = document.getElementById("miAutoWoOnIssues")?.checked !== false;

  if (!asset_id) return alert("Select an asset.");
  if (!inspection_date) return alert("Select inspection date.");
  if (requireEvidence && failedRows.length) {
    const missingNote = failedRows.find((r) => !String(r?.note || "").trim());
    if (missingNote) return alert(`Add a comment for failed item: ${missingNote.label}`);
    if (!evidenceFiles.length) return alert("Upload at least one evidence photo for failed checklist items.");
  }
  if (msg) msg.textContent = "Saving inspection...";
  try {
    const res = await fetch(`${API}/maintenance/inspections`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        asset_id,
        inspection_date,
        inspector_name,
        notes,
        inspection_type,
        machine_hours,
        checklist,
        required_parts,
        defect_component,
        defect_severity,
        defect_risk,
        recommended_action,
        evidence_required: requireEvidence ? 1 : 0,
        evidence_photo_count: evidenceFiles.length,
        enforce_evidence_rules: requireEvidence ? 1 : 0,
        create_work_order,
        create_work_order_on_issues,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save inspection");
    if (Number(data.id || 0) > 0 && evidenceFiles.length) {
      for (const file of evidenceFiles) {
        const fd = new FormData();
        fd.append("file", file);
        const up = await fetch(`${API}/maintenance/inspections/${Number(data.id)}/photo?caption=${encodeURIComponent("Failure evidence")}`, {
          method: "POST",
          body: fd,
        });
        const upData = await up.json().catch(() => ({}));
        if (!up.ok) throw new Error(upData.error || "Evidence photo upload failed");
      }
    }
    const wo = data.work_order_id ? ` Work order #${data.work_order_id} created.` : "";
    if (msg) msg.textContent = `Inspection saved.${wo}`;
    resetManagerInspectionForm();
    await loadManagerInspections();
  } catch (e) {
    if (msg) msg.textContent = `Save error: ${e.message || e}`;
  }
}

async function uploadInspectionPhoto(inspectionId) {
  const fileEl = document.getElementById(`miPhotoFile-${inspectionId}`);
  const capEl = document.getElementById(`miPhotoCaption-${inspectionId}`);
  const file = fileEl?.files?.[0];
  if (!file) return alert("Choose a photo first.");
  const fd = new FormData();
  fd.append("file", file);
  const caption = String(capEl?.value || "").trim();
  const q = caption ? `?caption=${encodeURIComponent(caption)}` : "";
  const res = await fetch(`${API}/maintenance/inspections/${inspectionId}/photo${q}`, {
    method: "POST",
    body: fd,
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Photo upload failed");
  await loadManagerInspections();
}

async function saveDamageReport() {
  const asset_id = Number(document.getElementById("drAsset")?.value || 0);
  const report_date = String(document.getElementById("drDate")?.value || "").trim();
  const damage_time = String(document.getElementById("drTime")?.value || "").trim();
  const inspector_name = String(document.getElementById("drInspector")?.value || "").trim();
  const hour_meter_raw = String(document.getElementById("drHours")?.value || "").trim();
  const damage_location = String(document.getElementById("drLocation")?.value || "").trim();
  const responsible_person = String(document.getElementById("drResponsiblePerson")?.value || "").trim();
  const severity = String(document.getElementById("drSeverity")?.value || "").trim();
  const damage_description = String(document.getElementById("drDescription")?.value || "").trim();
  const immediate_action = String(document.getElementById("drAction")?.value || "").trim();
  const out_of_service = document.getElementById("drOutOfService")?.checked ? 1 : 0;
  const pending_investigation = document.getElementById("drPendingInvestigation")?.checked ? 1 : 0;
  const hse_report_available = document.getElementById("drHseReportAvailable")?.checked ? 1 : 0;
  const msg = document.getElementById("drMsg");

  if (!asset_id) return alert("Select an asset for damage report.");
  if (!report_date) return alert("Select damage report date.");
  if (!damage_location) return alert("Enter damage location.");
  if (!severity) return alert("Select severity.");
  if (!damage_description) return alert("Enter damage description.");
  if (!immediate_action) return alert("Enter immediate action.");

  const hour_meter = hour_meter_raw === "" ? null : Number(hour_meter_raw);
  if (hour_meter != null && !Number.isFinite(hour_meter)) {
    return alert("Machine hours must be numeric.");
  }

  if (msg) msg.textContent = "Saving damage report...";
  try {
    const res = await fetch(`${API}/maintenance/damage-reports`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        asset_id,
        report_date,
        damage_time,
        inspector_name,
        hour_meter,
        damage_location,
        responsible_person,
        severity,
        damage_description,
        immediate_action,
        out_of_service,
        pending_investigation,
        hse_report_available,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save damage report");
    if (msg) msg.textContent = "Damage report saved.";
    document.getElementById("drDescription").value = "";
    document.getElementById("drAction").value = "";
    document.getElementById("drLocation").value = "";
    document.getElementById("drResponsiblePerson").value = "";
    document.getElementById("drHours").value = "";
    document.getElementById("drTime").value = "";
    document.getElementById("drSeverity").value = "";
    document.getElementById("drOutOfService").checked = false;
    document.getElementById("drPendingInvestigation").checked = false;
    document.getElementById("drHseReportAvailable").checked = false;
    await loadDamageReports();
  } catch (e) {
    if (msg) msg.textContent = `Save error: ${e.message || e}`;
  }
}

async function uploadDamagePhoto(reportId) {
  const fileEl = document.getElementById(`drPhotoFile-${reportId}`);
  const capEl = document.getElementById(`drPhotoCaption-${reportId}`);
  const file = fileEl?.files?.[0];
  if (!file) return alert("Choose a photo first.");
  const fd = new FormData();
  fd.append("file", file);
  const caption = String(capEl?.value || "").trim();
  const q = caption ? `?caption=${encodeURIComponent(caption)}` : "";
  const res = await fetch(`${API}/maintenance/damage-reports/${reportId}/photo${q}`, {
    method: "POST",
    body: fd,
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data.error || "Damage photo upload failed");
  await loadDamageReports();
}

function damageCard(r) {
  const photos = Array.isArray(r.photos) ? r.photos : [];
  const sev = String(r.severity || "").toLowerCase();
  const sevClass = sev === "critical" || sev === "high" ? "status-overdue" : sev === "medium" ? "status-soon" : "status-ok";
  const photoHtml = photos.length
    ? `<div class="row stack-10">${photos.map((p) => {
        const src = normImgSrc(p.file_path);
        return `<div style="display:flex; flex-direction:column; gap:4px;">
          <img src="${esc(src)}" alt="damage photo" style="width:140px; height:100px; object-fit:cover; border:1px solid #d1d5db; border-radius:8px;" />
          <small class="muted">${esc(p.caption || "")}</small>
        </div>`;
      }).join("")}</div>`
    : `<small class="muted">No photos yet.</small>`;

  return `
    <div class="card">
      <div><b>${esc(r.asset_code)}</b> - ${esc(r.asset_name || "")}</div>
      <div><small>Date: ${esc(r.report_date || "-")} ${esc(r.damage_time || "")} | Inspector: ${esc(r.inspector_name || "-")} | Hours: ${esc(r.hour_meter == null ? "-" : String(Number(r.hour_meter).toFixed(1)))}</small></div>
      <div><small>Location: <b>${esc(r.damage_location || "-")}</b> | Responsible: <b>${esc(r.responsible_person || "-")}</b></small></div>
      <div><small>Severity: <span class="${sevClass}">${esc((r.severity || "-").toUpperCase())}</span> | Out of service: <b>${Number(r.out_of_service || 0) ? "YES" : "NO"}</b> | Pending investigation: <b>${Number(r.pending_investigation || 0) ? "YES" : "NO"}</b> | HSE report: <b>${Number(r.hse_report_available || 0) ? "YES" : "NO"}</b></small></div>
      <div style="margin-top:6px;"><small><b>Damage:</b> ${esc(r.damage_description || "")}</small></div>
      <div style="margin-top:4px;"><small><b>Immediate action:</b> ${esc(r.immediate_action || "")}</small></div>
      <div style="margin-top:8px;">${photoHtml}</div>
      <div class="row stack-10" style="margin-top:8px;">
        <button data-dr-open-pdf="${Number(r.id)}">Open PDF</button>
        <button data-dr-download-pdf="${Number(r.id)}">Download PDF</button>
        <input id="drPhotoFile-${Number(r.id)}" type="file" accept="image/*" />
        <input id="drPhotoCaption-${Number(r.id)}" class="w-200" placeholder="Photo caption (optional)" />
        <button data-dr-upload="${Number(r.id)}">Upload Photo</button>
      </div>
    </div>
  `;
}

async function loadDamageReports() {
  const list = document.getElementById("drList");
  if (!list) return;
  list.innerHTML = `<div class="skeleton-block"></div>`;
  const asset_id = String(document.getElementById("drFilterAsset")?.value || "").trim();
  const start = String(document.getElementById("drStart")?.value || "").trim();
  const end = String(document.getElementById("drEnd")?.value || "").trim();
  const responsible_person = String(document.getElementById("drFilterResponsiblePerson")?.value || "").trim();
  const pending_investigation = String(document.getElementById("drFilterPendingInvestigation")?.value || "").trim();
  const hse_report_available = String(document.getElementById("drFilterHseReportAvailable")?.value || "").trim();
  const q = new URLSearchParams();
  if (asset_id) q.set("asset_id", asset_id);
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  if (responsible_person) q.set("responsible_person", responsible_person);
  if (pending_investigation === "0" || pending_investigation === "1") q.set("pending_investigation", pending_investigation);
  if (hse_report_available === "0" || hse_report_available === "1") q.set("hse_report_available", hse_report_available);
  try {
    const res = await fetch(`${API}/maintenance/damage-reports${q.toString() ? `?${q.toString()}` : ""}`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load damage reports");
    const rows = Array.isArray(data.rows) ? data.rows : [];
    list.innerHTML = rows.length ? rows.map(damageCard).join("") : `<div class="muted">No damage reports found.</div>`;
  } catch (e) {
    list.innerHTML = `<div class="message-error">Damage report load error: ${esc(e.message || e)}</div>`;
  }
}

function openDamageReportPdf(id, download = false) {
  const n = Number(id || 0);
  if (!n) return;
  const q = download ? "?download=1" : "";
  return openProtectedPdf(`${API}/reports/damage-report/${n}.pdf${q}`, {
    download,
    filename: `IRONLOG_Damage_Report_${n}.pdf`,
  });
}

function openDamageReportsBulkPdf(download = false, withPhotos = false) {
  const start = String(document.getElementById("drStart")?.value || "").trim();
  const end = String(document.getElementById("drEnd")?.value || "").trim();
  const assetId = String(document.getElementById("drFilterAsset")?.value || "").trim();
  if (!start || !end) {
    alert("Select damage report start and end dates first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (assetId) q.set("asset_id", assetId);
  if (withPhotos) q.set("with_photos", "1");
  if (download) q.set("download", "1");
  return openProtectedPdf(`${API}/reports/damage-reports.pdf?${q.toString()}`, {
    download,
    filename: `IRONLOG_Damage_Reports_${start}_to_${end}.pdf`,
  });
}

function downloadDamageReportsXlsx() {
  const start = String(document.getElementById("drStart")?.value || "").trim();
  const end = String(document.getElementById("drEnd")?.value || "").trim();
  const assetId = String(document.getElementById("drFilterAsset")?.value || "").trim();
  if (!start || !end) {
    alert("Select damage report start and end dates first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (assetId) q.set("asset_id", assetId);
  return downloadProtectedXlsxFile(
    `${API}/reports/damage-reports.xlsx?${q.toString()}`,
    `IRONLOG_Damage_Reports_${start}_to_${end}.xlsx`,
  ).catch((err) => alert(`Could not download damage report export: ${err.message || err}`));
}

function openManagerInspectionPdf(id, download = false) {
  const n = Number(id || 0);
  if (!n) return;
  const q = download ? "?download=1" : "";
  return openProtectedPdf(`${API}/reports/manager-inspection/${n}.pdf${q}`, {
    download,
    filename: `IRONLOG_Manager_Inspection_${n}.pdf`,
  });
}

function openManagerInspectionsBulkPdf(download = false, withPhotos = false) {
  const start = String(document.getElementById("miStart")?.value || "").trim();
  const end = String(document.getElementById("miEnd")?.value || "").trim();
  const assetId = String(document.getElementById("miFilterAsset")?.value || "").trim();
  if (!start || !end) {
    alert("Select start and end dates first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (assetId) q.set("asset_id", assetId);
  if (withPhotos) q.set("with_photos", "1");
  if (download) q.set("download", "1");
  return openProtectedPdf(`${API}/reports/manager-inspections.pdf?${q.toString()}`, {
    download,
    filename: `IRONLOG_Manager_Inspections_${start}_to_${end}.pdf`,
  });
}

function inspectionCard(r) {
  const formatComponentNotesHtml = (raw) => {
    const text = String(raw || "").trim();
    if (!text) return "<small>-</small>";
    const rows = text
      .split(/\r?\n|;/)
      .map((s) => String(s || "").trim())
      .filter(Boolean)
      .map((line) => {
        const m = line.match(/^([^:|-]+)\s*[:|-]\s*(.+)$/);
        if (m) return { component: String(m[1] || "").trim(), note: String(m[2] || "").trim() };
        return { component: "General", note: line };
      });
    if (!rows.length) return "<small>-</small>";
    return rows
      .map((n) => `<small><b>${esc(n.component)}:</b> ${esc(n.note)}</small>`)
      .join("<br/>");
  };

  const photos = Array.isArray(r.photos) ? r.photos : [];
  const photoHtml = photos.length
    ? `<div class="row stack-10">${photos.map((p) => {
        const src = normImgSrc(p.file_path);
        return `<div style="display:flex; flex-direction:column; gap:4px;">
          <img src="${esc(src)}" alt="inspection photo" style="width:140px; height:100px; object-fit:cover; border:1px solid #d1d5db; border-radius:8px;" />
          <small class="muted">${esc(p.caption || "")}</small>
        </div>`;
      }).join("")}</div>`
    : `<small class="muted">No photos yet.</small>`;

  const hrs =
    r.machine_hours != null && Number.isFinite(Number(r.machine_hours))
      ? Number(r.machine_hours).toFixed(1)
      : "—";
  const live =
    r.live_hours_snapshot != null && Number.isFinite(Number(r.live_hours_snapshot))
      ? `${Number(r.live_hours_snapshot).toFixed(1)} (${esc(r.live_hours_source || "—")})`
      : "—";
  const wo =
    r.work_order_id != null && Number(r.work_order_id) > 0
      ? `<b>WO #${Number(r.work_order_id)}</b>`
      : `<span class="muted">No WO</span>`;
  const chk = Array.isArray(r.checklist) ? r.checklist : [];
  const fails = chk.filter((c) => c.ok === false);
  const failLine = fails.length
    ? `<div style="margin-top:4px;"><small class="status-overdue">Checklist fail: ${fails.map((c) => esc(c.label || c.key)).join("; ")}</small></div>`
    : "";
  const parts = Array.isArray(r.required_parts) ? r.required_parts : [];
  const partsLine = parts.length
    ? `<div style="margin-top:4px;"><small><b>Parts:</b> ${parts.map((p) => `${esc(p.part_code)} × ${esc(String(p.qty))}`).join(", ")}</small></div>`
    : "";

  return `
    <div class="card">
      <div><b>${esc(r.asset_code)}</b> - ${esc(r.asset_name || "")}</div>
      <div><small>Date: ${esc(r.inspection_date)} | Inspector: ${esc(r.inspector_name || "-")}</small></div>
      <div><small>Type: <b>${esc(String(r.inspection_type || "machine_general").replaceAll("_", " "))}</b></small></div>
      <div><small>Machine hrs: ${esc(hrs)} | Live snapshot: ${live} | ${wo}</small></div>
      ${failLine}
      ${partsLine}
      <div style="margin-top:6px;">${formatComponentNotesHtml(r.notes)}</div>
      <div style="margin-top:8px;">${photoHtml}</div>
      <div class="row stack-10" style="margin-top:8px;">
        <button data-mi-open-pdf="${Number(r.id)}">Open PDF</button>
        <button data-mi-download-pdf="${Number(r.id)}">Download PDF</button>
        <button data-mi-delete="${Number(r.id)}">Delete</button>
        <input id="miPhotoFile-${Number(r.id)}" type="file" accept="image/*" />
        <input id="miPhotoCaption-${Number(r.id)}" class="w-200" placeholder="Photo caption (optional)" />
        <button data-mi-upload="${Number(r.id)}">Upload Photo</button>
      </div>
    </div>
  `;
}

async function loadManagerInspections() {
  const list = document.getElementById("miList");
  if (!list) return;
  list.innerHTML = `<div class="skeleton-block"></div>`;
  const asset_id = String(document.getElementById("miFilterAsset")?.value || "").trim();
  const start = String(document.getElementById("miStart")?.value || "").trim();
  const end = String(document.getElementById("miEnd")?.value || "").trim();
  const q = new URLSearchParams();
  if (asset_id) q.set("asset_id", asset_id);
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  try {
    const res = await fetch(`${API}/maintenance/inspections${q.toString() ? `?${q.toString()}` : ""}`, {
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load inspections");
    const rows = Array.isArray(data.rows)
      ? data.rows
      : (Array.isArray(data.data) ? data.data : []);
    list.innerHTML = rows.length ? rows.map(inspectionCard).join("") : `<div class="muted">No inspections found.</div>`;
  } catch (e) {
    list.innerHTML = `<div class="message-error">Inspection load error: ${esc(e.message || e)}</div>`;
  }
}

async function deleteManagerInspection(inspectionId) {
  const id = Number(inspectionId || 0);
  if (!id) return;
  const ok = window.confirm(`Delete manager inspection #${id}? This cannot be undone.`);
  if (!ok) return;
  const res = await fetch(`${API}/maintenance/inspections/${id}`, {
    method: "DELETE",
    headers: authHeaders(),
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(data.error || "Failed to delete inspection");
  await loadManagerInspections();
}

function getMiPartByCode(codeIn) {
  return getWfPartByCode(codeIn);
}

const MANAGER_INSPECTION_CHECKLISTS = {
  machine_general: [
    { key: "structure", label: "Structure / visible damage" },
    { key: "fluids", label: "Fluids / leaks" },
    { key: "tyres", label: "Tyres / undercarriage" },
    { key: "safety", label: "Safety & access (rails, steps, extinguisher)" },
    { key: "cabin", label: "Cabin / visibility / instruments" },
    { key: "attachments", label: "GET / tools / attachments" },
    { key: "housekeeping", label: "Housekeeping" },
    { key: "noise", label: "Operation / unusual noise" },
  ],
  environmental: [
    { key: "spill_control", label: "Spill control and containment in place" },
    { key: "waste_segregation", label: "Waste segregation and disposal compliant" },
    { key: "oil_storage", label: "Fuel/oil/chemical storage condition acceptable" },
    { key: "drainage", label: "Drainage channels and runoff protection clear" },
    { key: "housekeeping_env", label: "Work area free from environmental hazards" },
  ],
  fire_extinguishers: [
    { key: "ext_present", label: "Extinguisher present and accessible" },
    { key: "ext_seal_pin", label: "Safety pin and tamper seal intact" },
    { key: "ext_pressure", label: "Pressure gauge in serviceable range" },
    { key: "ext_label", label: "Inspection tag and expiry date valid" },
    { key: "ext_mounting", label: "Mounting bracket/signage condition acceptable" },
  ],
  hand_tools: [
    { key: "tool_condition", label: "No cracked, bent, or damaged hand tools" },
    { key: "tool_storage", label: "Correct storage and housekeeping for tools" },
    { key: "tool_tracking", label: "Tool control / inventory status updated" },
    { key: "tool_safety", label: "Safe use practices observed on site" },
  ],
  electrical_tools_equipment: [
    { key: "cable_plug", label: "Cables/plugs free from damage" },
    { key: "isolation_switch", label: "Isolation switch/lockout available and used" },
    { key: "earth_leakage", label: "Earth leakage / test status compliant" },
    { key: "guarding", label: "Machine/tool guards fitted and functional" },
    { key: "electrical_ppe", label: "Electrical PPE and safe operation followed" },
  ],
};
function getManagerInspectionType() {
  return String(document.getElementById("miInspectionType")?.value || "machine_general").trim();
}
function getManagerInspectionChecklistTemplate() {
  const type = getManagerInspectionType();
  return MANAGER_INSPECTION_CHECKLISTS[type] || MANAGER_INSPECTION_CHECKLISTS.machine_general;
}

function renderManagerInspectionChecklist() {
  const host = document.getElementById("miChecklist");
  if (!host) return;
  const template = getManagerInspectionChecklistTemplate();
  host.innerHTML = template.map(
    (row) => `
    <div class="row stack-10" style="align-items:center; flex-wrap:wrap; gap:8px;">
      <span style="min-width:240px; font-size:13px;">${esc(row.label)}</span>
      <span class="row stack-10" style="gap:10px;">
        <label><input type="radio" name="miChk-${esc(row.key)}" value="ok" /> OK</label>
        <label><input type="radio" name="miChk-${esc(row.key)}" value="fail" /> Fail</label>
        <label><input type="radio" name="miChk-${esc(row.key)}" value="na" /> N/A</label>
      </span>
      <input type="text" class="w-200 mi-chk-note" data-mi-chk="${esc(row.key)}" placeholder="Note (optional)" />
    </div>`
  ).join("");
}

function collectManagerInspectionChecklist() {
  const template = getManagerInspectionChecklistTemplate();
  return template.map((row) => {
    const sel = document.querySelector(`input[name="miChk-${row.key}"]:checked`);
    const val = sel ? String(sel.value || "") : "";
    let ok = null;
    if (val === "ok") ok = true;
    else if (val === "fail") ok = false;
    const noteEl = document.querySelector(`input.mi-chk-note[data-mi-chk="${row.key}"]`);
    const note = String(noteEl?.value || "").trim() || null;
    return { key: row.key, label: row.label, ok, note };
  });
}

function getFailedChecklistRows(checklist) {
  return (Array.isArray(checklist) ? checklist : []).filter((c) => c && c.ok === false);
}

function appendVoiceToField(fieldId) {
  const out = document.getElementById(fieldId);
  if (!out) return;
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    alert("Voice dictation is not supported in this browser.");
    return;
  }
  const rec = new SpeechRecognition();
  rec.lang = "en-US";
  rec.interimResults = false;
  rec.maxAlternatives = 1;
  rec.onresult = (evt) => {
    const text = String(evt?.results?.[0]?.[0]?.transcript || "").trim();
    if (!text) return;
    const existing = String(out.value || "").trim();
    out.value = existing ? `${existing} ${text}` : text;
    out.dispatchEvent(new Event("input", { bubbles: true }));
  };
  rec.onerror = () => {
    alert("Voice dictation failed. Please try again.");
  };
  rec.start();
}

function getSelectedAssetCode(selectId) {
  const sel = document.getElementById(selectId);
  if (!sel) return "";
  const opt = sel.options?.[sel.selectedIndex];
  if (!opt) return "";
  const label = String(opt.textContent || "").trim();
  const code = label.split("-")[0]?.trim();
  return String(code || "").toUpperCase();
}

function resolveScanAssetCode() {
  const scan = String(document.getElementById("scanAssetCode")?.value || "").trim().toUpperCase();
  if (scan) return scan;
  return getSelectedAssetCode("miAsset");
}

function addManagerInspectionPartRow(partCode = "", qty = "", note = "") {
  const body = document.getElementById("miPartsBody");
  if (!body) return;
  const tr = document.createElement("tr");
  tr.innerHTML = `
    <td><input class="mi-part-code" type="text" list="miPartsList" style="min-width:160px;" value="${esc(partCode)}" /></td>
    <td><input class="mi-part-qty" type="number" step="0.01" min="0" style="width:90px;" value="${esc(qty)}" /></td>
    <td><input class="mi-part-note" type="text" style="min-width:140px;" value="${esc(note)}" /></td>
    <td><button type="button" class="mi-part-remove">Remove</button></td>`;
  body.appendChild(tr);
  tr.querySelector(".mi-part-remove")?.addEventListener("click", () => {
    tr.remove();
  });
}

function collectManagerInspectionParts() {
  const body = document.getElementById("miPartsBody");
  if (!body) return [];
  const out = [];
  body.querySelectorAll("tr").forEach((tr) => {
    const part_code = String(tr.querySelector(".mi-part-code")?.value || "").trim();
    const qty = Number(tr.querySelector(".mi-part-qty")?.value || 0);
    const note = String(tr.querySelector(".mi-part-note")?.value || "").trim() || null;
    if (!part_code || !Number.isFinite(qty) || qty <= 0) return;
    const meta = getMiPartByCode(part_code);
    out.push({
      part_id: meta?.id != null ? Number(meta.id) : null,
      part_code,
      qty,
      note,
    });
  });
  return out;
}

async function pullManagerInspectionLiveHours() {
  const assetId = Number(document.getElementById("miAsset")?.value || 0);
  const inspectionDate = String(document.getElementById("miDate")?.value || "").trim();
  const meta = document.getElementById("miLiveMeta");
  const inp = document.getElementById("miMachineHours");
  if (!assetId) {
    if (meta) meta.textContent = "Select an asset first.";
    return;
  }
  if (meta) meta.textContent = "Loading live hours…";
  try {
    const q = new URLSearchParams();
    if (/^\d{4}-\d{2}-\d{2}$/.test(inspectionDate)) q.set("as_of", inspectionDate);
    const qs = q.toString();
    const url = `${API}/maintenance/asset/${assetId}/live-hours${qs ? `?${qs}` : ""}`;
    const res = await fetch(url, { headers: authHeaders() });
    let data;
    try {
      data = await res.json();
    } catch {
      throw new Error(`Live hours response was not JSON (HTTP ${res.status}).`);
    }
    if (!res.ok) throw new Error(data?.error || "Failed to load live hours");
    const h = Number(data.current_hours ?? 0);
    const src = String(data.current_hours_source || "");
    const asOf = data.as_of ? ` up to ${data.as_of}` : "";
    if (inp) inp.value = Number.isFinite(h) ? String(Number(h).toFixed(1)) : "";
    const srcLabel =
      {
        daily_closing: "Daily closing",
        daily_sum: "Daily sum",
        asset_hours: "Asset hours",
      }[src] || src || "—";
    let line = `Live${asOf}: ${Number.isFinite(h) ? h.toFixed(1) : "—"} h (${srcLabel})`;
    if (Number.isFinite(h) && h <= 0) {
      line +=
        " — No meter or usage found for this asset/date; add Daily Input or enter hours manually.";
    }
    if (meta) meta.textContent = line;
  } catch (e) {
    console.error("pullManagerInspectionLiveHours:", e);
    if (meta) meta.textContent = e.message || String(e);
  }
}

function resetManagerInspectionForm() {
  document.getElementById("miNotes").value = "";
  const miDefectComponent = document.getElementById("miDefectComponent");
  if (miDefectComponent) miDefectComponent.value = "";
  const miDefectSeverity = document.getElementById("miDefectSeverity");
  if (miDefectSeverity) miDefectSeverity.value = "";
  const miDefectRisk = document.getElementById("miDefectRisk");
  if (miDefectRisk) miDefectRisk.value = "";
  const miRecommendedAction = document.getElementById("miRecommendedAction");
  if (miRecommendedAction) miRecommendedAction.value = "";
  const miRequireEvidence = document.getElementById("miRequireEvidence");
  if (miRequireEvidence) miRequireEvidence.checked = true;
  const miEvidencePhotos = document.getElementById("miEvidencePhotos");
  if (miEvidencePhotos) miEvidencePhotos.value = "";
  const template = getManagerInspectionChecklistTemplate();
  template.forEach((row) => {
    document.querySelectorAll(`input[name="miChk-${row.key}"]`).forEach((r) => {
      r.checked = false;
    });
    const ne = document.querySelector(`input.mi-chk-note[data-mi-chk="${row.key}"]`);
    if (ne) ne.value = "";
  });
  const body = document.getElementById("miPartsBody");
  if (body) {
    body.innerHTML = "";
    addManagerInspectionPartRow();
  }
  const auto = document.getElementById("miAutoWoOnIssues");
  if (auto) auto.checked = true;
  const al = document.getElementById("miCreateWoAlways");
  if (al) al.checked = false;
  const lm = document.getElementById("miLiveMeta");
  if (lm) lm.textContent = "";
}

const ARTISAN_INSPECTION_CHECKLIST = [
  { key: "prestart", label: "Pre-start visual condition (machine / plant)" },
  { key: "guards", label: "Guards, covers, and safety devices" },
  { key: "hydraulics", label: "Hydraulic hoses, leaks, and fittings" },
  { key: "electrical", label: "Electrical panels / cabling / lights" },
  { key: "lubrication", label: "Lubrication points / levels" },
  { key: "brakes_steering", label: "Brakes / steering / controls response" },
  { key: "alarms", label: "Alarms, horn, and warning systems" },
  { key: "housekeeping", label: "Housekeeping around machine / plant" },
];

function renderArtisanInspectionChecklist() {
  const host = document.getElementById("aiChecklist");
  if (!host) return;
  host.innerHTML = ARTISAN_INSPECTION_CHECKLIST.map(
    (row) => `
    <div class="row stack-10" style="align-items:center; flex-wrap:wrap; gap:8px;">
      <span style="min-width:240px; font-size:13px;">${esc(row.label)}</span>
      <span class="row stack-10" style="gap:10px;">
        <label><input type="radio" name="aiChk-${esc(row.key)}" value="ok" /> OK</label>
        <label><input type="radio" name="aiChk-${esc(row.key)}" value="fail" /> Fail</label>
        <label><input type="radio" name="aiChk-${esc(row.key)}" value="na" /> N/A</label>
      </span>
      <input type="text" class="w-200 ai-chk-note" data-ai-chk="${esc(row.key)}" placeholder="Note (optional)" />
    </div>`
  ).join("");
}

function collectArtisanInspectionChecklist() {
  return ARTISAN_INSPECTION_CHECKLIST.map((row) => {
    const sel = document.querySelector(`input[name="aiChk-${row.key}"]:checked`);
    const val = sel ? String(sel.value || "") : "";
    let ok = null;
    if (val === "ok") ok = true;
    else if (val === "fail") ok = false;
    const noteEl = document.querySelector(`input.ai-chk-note[data-ai-chk="${row.key}"]`);
    const note = String(noteEl?.value || "").trim() || null;
    return { key: row.key, label: row.label, ok, note };
  });
}

async function pullArtisanInspectionLiveHours() {
  const assetId = Number(document.getElementById("aiAsset")?.value || 0);
  const inspectionDate = String(document.getElementById("aiDate")?.value || "").trim();
  const meta = document.getElementById("aiLiveMeta");
  const inp = document.getElementById("aiMachineHours");
  if (!assetId) {
    if (meta) meta.textContent = "Select an asset first.";
    return;
  }
  if (meta) meta.textContent = "Loading live hours…";
  try {
    const q = new URLSearchParams();
    if (/^\d{4}-\d{2}-\d{2}$/.test(inspectionDate)) q.set("as_of", inspectionDate);
    const qs = q.toString();
    const url = `${API}/maintenance/asset/${assetId}/live-hours${qs ? `?${qs}` : ""}`;
    const res = await fetch(url, { headers: authHeaders() });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Failed to load live hours");
    const h = Number(data.current_hours ?? 0);
    const src = String(data.current_hours_source || "");
    const asOf = data.as_of ? ` up to ${data.as_of}` : "";
    if (inp) inp.value = Number.isFinite(h) ? String(Number(h).toFixed(1)) : "";
    const srcLabel = {
      daily_closing: "Daily closing",
      daily_sum: "Daily sum",
      asset_hours: "Asset hours",
    }[src] || src || "—";
    if (meta) meta.textContent = `Live${asOf}: ${Number.isFinite(h) ? h.toFixed(1) : "—"} h (${srcLabel})`;
  } catch (e) {
    if (meta) meta.textContent = e.message || String(e);
  }
}

function resetArtisanInspectionForm() {
  const notes = document.getElementById("aiNotes");
  if (notes) notes.value = "";
  const shift = document.getElementById("aiShift");
  if (shift) shift.value = "";
  const formNo = document.getElementById("aiFormNumber");
  if (formNo) formNo.value = generateArtisanFormNumber();
  const live = document.getElementById("aiLiveMeta");
  if (live) live.textContent = "";
  ARTISAN_INSPECTION_CHECKLIST.forEach((row) => {
    document.querySelectorAll(`input[name="aiChk-${row.key}"]`).forEach((r) => {
      r.checked = false;
    });
    const ne = document.querySelector(`input.ai-chk-note[data-ai-chk="${row.key}"]`);
    if (ne) ne.value = "";
  });
}

function generateArtisanFormNumber() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}`;
  const rand = Math.floor(Math.random() * 900 + 100);
  return `AI-${stamp}-${rand}`;
}

async function saveArtisanInspection() {
  const asset_id = Number(document.getElementById("aiAsset")?.value || 0);
  const inspection_date = String(document.getElementById("aiDate")?.value || "").trim();
  const inspector_name = String(document.getElementById("aiInspector")?.value || "").trim();
  const shift = String(document.getElementById("aiShift")?.value || "").trim().toLowerCase();
  const form_number = String(document.getElementById("aiFormNumber")?.value || "").trim();
  const notes = String(document.getElementById("aiNotes")?.value || "").trim();
  const msg = document.getElementById("aiMsg");
  const mhRaw = String(document.getElementById("aiMachineHours")?.value || "").trim();
  const machine_hours = mhRaw === "" ? null : Number(mhRaw);
  if (machine_hours != null && !Number.isFinite(machine_hours)) {
    return alert("Machine hours must be a number.");
  }
  const checklist = collectArtisanInspectionChecklist();

  if (!asset_id) return alert("Select an asset.");
  if (!inspection_date) return alert("Select inspection date.");
  if (msg) msg.textContent = "Saving artisan inspection...";
  try {
    const res = await fetch(`${API}/maintenance/artisan-inspections`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        asset_id,
        inspection_date,
        inspector_name,
        shift: shift || null,
        form_number: form_number || null,
        notes,
        machine_hours,
        checklist,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save artisan inspection");
    if (msg) msg.textContent = "Artisan inspection saved.";
    resetArtisanInspectionForm();
    await loadArtisanInspections();
  } catch (e) {
    if (msg) msg.textContent = `Save error: ${e.message || e}`;
  }
}

function openArtisanInspectionPdf(id, download = false) {
  const n = Number(id || 0);
  if (!n) return;
  const q = download ? "?download=1" : "";
  return openProtectedPdf(`${API}/reports/artisan-inspection/${n}.pdf${q}`, {
    download,
    filename: `IRONLOG_Artisan_Inspection_${n}.pdf`,
  });
}

const ARTISAN_RESULT_TEXT = { fit: "Fit for work", restricted: "Fit with restrictions", not_fit: "Not fit for work" };

/** Details added by inspections done in the technician portal. */
function artisanExtraLine(r) {
  const bits = [];
  if (r.inspection_type) bits.push(`Type: ${esc(String(r.inspection_type).replace(/_/g, " "))}`);
  if (r.location) bits.push(`Location: ${esc(r.location)}`);
  if (r.overall_result) bits.push(`Result: <b class="${r.overall_result === "not_fit" ? "status-overdue" : ""}">${esc(ARTISAN_RESULT_TEXT[r.overall_result] || r.overall_result)}</b>`);
  if (Number(r.photo_count) > 0) bits.push(`${Number(r.photo_count)} photo${Number(r.photo_count) === 1 ? "" : "s"}`);
  if (r.work_order_id) bits.push(`<a href="./workorders.html?wo=${Number(r.work_order_id)}">WO #${Number(r.work_order_id)}</a>`);
  return bits.length ? `<div><small>${bits.join(" | ")}</small></div>` : "";
}

function artisanInspectionCard(r) {
  const hrs = r.machine_hours != null && Number.isFinite(Number(r.machine_hours))
    ? Number(r.machine_hours).toFixed(1)
    : "—";
  const live = r.live_hours_snapshot != null && Number.isFinite(Number(r.live_hours_snapshot))
    ? `${Number(r.live_hours_snapshot).toFixed(1)} (${esc(r.live_hours_source || "—")})`
    : "—";
  const shift = String(r.shift || "").trim();
  const formNo = String(r.form_number || "").trim();
  const chk = Array.isArray(r.checklist) ? r.checklist : [];
  const fails = chk.filter((c) => c.ok === false);
  const failLine = fails.length
    ? `<div style="margin-top:4px;"><small class="status-overdue">Checklist fail: ${fails.map((c) => esc(c.label || c.key)).join("; ")}</small></div>`
    : "";

  return `
    <div class="card">
      <div><b>${esc(r.asset_code)}</b> - ${esc(r.asset_name || "")}</div>
      <div><small>Date: ${esc(r.inspection_date)}${shift ? ` | Shift: ${esc(shift.toUpperCase())}` : ""} | Artisan: ${esc(r.inspector_name || "-")}</small></div>
      <div><small>Form No: <b>${esc(formNo || "-")}</b></small></div>
      <div><small>Machine hrs: ${esc(hrs)} | Live snapshot: ${live}</small></div>
      ${artisanExtraLine(r)}
      ${failLine}
      <div style="margin-top:6px;"><small>${esc(r.notes || "")}</small></div>
      <div class="row stack-10" style="margin-top:8px;">
        <button data-ai-open-pdf="${Number(r.id)}">Open PDF</button>
        <button data-ai-download-pdf="${Number(r.id)}">Download PDF</button>
      </div>
    </div>
  `;
}

function openArtisanBlankFormPdf(download = false) {
  const date = String(document.getElementById("aiDate")?.value || "").trim();
  const assetId = String(document.getElementById("aiAsset")?.value || "").trim();
  const shift = String(document.getElementById("aiShift")?.value || "").trim();
  const inspector = String(document.getElementById("aiInspector")?.value || "").trim();
  const formNoEl = document.getElementById("aiFormNumber");
  if (formNoEl && !String(formNoEl.value || "").trim()) formNoEl.value = generateArtisanFormNumber();
  const formNo = String(formNoEl?.value || "").trim();
  const q = new URLSearchParams();
  if (date) q.set("date", date);
  if (assetId) q.set("asset_id", assetId);
  if (shift) q.set("shift", shift);
  if (inspector) q.set("inspector_name", inspector);
  if (formNo) q.set("form_number", formNo);
  if (download) q.set("download", "1");
  return openProtectedPdf(`${API}/reports/artisan-inspection-form.pdf?${q.toString()}`, {
    download,
    filename: `IRONLOG_Artisan_Inspection_Blank_${date || "form"}.pdf`,
  });
}

async function loadArtisanInspections() {
  const list = document.getElementById("aiList");
  if (!list) return;
  list.innerHTML = `<div class="skeleton-block"></div>`;
  const asset_id = String(document.getElementById("aiFilterAsset")?.value || "").trim();
  const start = String(document.getElementById("aiStart")?.value || "").trim();
  const end = String(document.getElementById("aiEnd")?.value || "").trim();
  const q = new URLSearchParams();
  if (asset_id) q.set("asset_id", asset_id);
  if (start) q.set("start", start);
  if (end) q.set("end", end);
  try {
    const res = await fetch(`${API}/maintenance/artisan-inspections${q.toString() ? `?${q.toString()}` : ""}`, {
      headers: authHeaders(),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load artisan inspections");
    const rows = Array.isArray(data.rows)
      ? data.rows
      : (Array.isArray(data.data) ? data.data : []);
    list.innerHTML = rows.length ? rows.map(artisanInspectionCard).join("") : `<div class="muted">No artisan inspections found.</div>`;
  } catch (e) {
    list.innerHTML = `<div class="message-error">Artisan inspection load error: ${esc(e.message || e)}</div>`;
  }
}

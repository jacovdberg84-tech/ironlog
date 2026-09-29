// IRONLOG/web/app/admin.js — Users & Access, safety/telematics admin, master data, SMTP, push, PDF settings, backups.
// Part of the main app; index.html loads these files in order and they share one global scope.

let __adminTabKeysLoaded = false;

const ADMIN_ROLE_OPTIONS = [
  { value: "admin", label: "Admin — unrestricted system access" },
  { value: "plant_manager", label: "Plant Manager — fleet, operations and reports" },
  { value: "workshop_admin", label: "Workshop Admin — maintenance and workshop control" },
  { value: "storeman", label: "Stores — stock, parts and receiving" },
  { value: "plant_clerk", label: "Plant Clerk — daily operational data capture" },
];

function populateAdminRolesSelect(rolesSel, selectedValues = null) {
  if (!rolesSel) return;
  const selected = new Set(
    (selectedValues != null
      ? selectedValues
      : Array.from(rolesSel.selectedOptions || []).map((o) => o.value)
    ).map((v) => String(v || "").trim().toLowerCase()).filter(Boolean)
  );
  rolesSel.innerHTML = "";
  ADMIN_ROLE_OPTIONS.forEach(({ value, label }) => {
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = label;
    opt.selected = selected.has(value);
    rolesSel.appendChild(opt);
  });
  if (!selected.size) {
    const op = rolesSel.querySelector('option[value="plant_clerk"]');
    if (op) op.selected = true;
  }
}

async function ensureAdminTabOptions() {
  const sel = qs("adminUserTabs");
  const rolesSel = qs("adminRoles");
  if (!sel && !rolesSel) return;

  if (rolesSel && rolesSel.options.length <= 1) {
    populateAdminRolesSelect(rolesSel);
  }

  if (__adminTabKeysLoaded) return;
  try {
    const data = await fetchJson(`${API}/api/auth/tabs`);
    const keys = Array.isArray(data.keys) ? data.keys : [];
    if (sel) {
      sel.innerHTML = "";
      keys.forEach((k) => {
        const opt = document.createElement("option");
        opt.value = k;
        opt.textContent = k;
        sel.appendChild(opt);
      });
    }
    const apiRoles = Array.isArray(data.roles) ? data.roles.map((r) => String(r || "").trim().toLowerCase()).filter(Boolean) : [];
    if (rolesSel && apiRoles.length) {
      const known = new Set(ADMIN_ROLE_OPTIONS.map((r) => r.value));
      const selectedBefore = Array.from(rolesSel.selectedOptions || []).map((o) => o.value);
      rolesSel.innerHTML = "";
      apiRoles.forEach((r) => {
        const meta = ADMIN_ROLE_OPTIONS.find((x) => x.value === r);
        const opt = document.createElement("option");
        opt.value = r;
        opt.textContent = meta ? meta.label : r;
        if (selectedBefore.includes(r) || (!selectedBefore.length && r === "operator")) opt.selected = true;
        rolesSel.appendChild(opt);
      });
      apiRoles.filter((r) => !known.has(r)).forEach((r) => {
        const opt = document.createElement("option");
        opt.value = r;
        opt.textContent = r;
        rolesSel.appendChild(opt);
      });
    } else if (rolesSel) {
      populateAdminRolesSelect(rolesSel);
    }
    __adminTabKeysLoaded = true;
  } catch {
    if (rolesSel && rolesSel.options.length <= 1) populateAdminRolesSelect(rolesSel);
  }
}

async function loadAdminUsers() {
  const pre = qs("adminUsersResult");
  const tbody = qs("adminUsersTbody");
  if (!tbody) return;
  setStatus("Loading users…");
  try {
    await ensureAdminTabOptions();
    const data = await fetchJson(`${API}/api/auth/users`);
    const rows = Array.isArray(data.rows) ? data.rows : [];
    const knownLocations = new Set(["main"]);
    rows.forEach((r) => {
      if (Array.isArray(r.allowed_locations)) {
        r.allowed_locations.forEach((loc) => {
          const v = String(loc || "").trim().toLowerCase();
          if (v) knownLocations.add(v);
        });
      }
    });
    const locList = qs("adminAllowedLocationsList");
    if (locList) {
      locList.innerHTML = "";
      Array.from(knownLocations).sort().forEach((loc) => {
        const opt = document.createElement("option");
        opt.value = loc;
        locList.appendChild(opt);
      });
    }
    tbody.innerHTML = "";
    rows.forEach((r) => {
      const tr = document.createElement("tr");
      const rolesText = Array.isArray(r.roles) && r.roles.length ? r.roles.join(", ") : String(r.role || "operator");
      const locText = Array.isArray(r.allowed_locations) && r.allowed_locations.length ? r.allowed_locations.join(", ") : "all";
      tr.innerHTML = `<td>${escapeHtml(r.username)}</td><td>${escapeHtml(r.full_name || "")}</td><td>${escapeHtml(r.department || "")}</td><td>${escapeHtml(rolesText)}</td><td>${escapeHtml(locText)}</td><td>${r.active ? "yes" : "no"}</td><td>${r.has_password ? "yes" : "no"}</td><td>${r.has_pin ? "yes" : "no"}</td>`;
      tr.style.cursor = "pointer";
      tr.addEventListener("click", () => {
        if (qs("adminUsername")) qs("adminUsername").value = r.username;
        if (qs("adminFullName")) qs("adminFullName").value = r.full_name || "";
        if (qs("adminDepartment")) qs("adminDepartment").value = r.department || "";
        if (qs("adminAllowedLocations")) {
          const loc = Array.isArray(r.allowed_locations) ? r.allowed_locations.join(",") : "";
          qs("adminAllowedLocations").value = loc;
        }
        const rolesSel = qs("adminRoles");
        if (rolesSel) {
          const selectedRoles = Array.isArray(r.roles) && r.roles.length ? r.roles : [String(r.role || "operator")];
          Array.from(rolesSel.options).forEach((o) => {
            o.selected = selectedRoles.includes(o.value);
          });
        }
        if (qs("adminPassword")) qs("adminPassword").value = "";
        if (qs("adminPin")) qs("adminPin").value = "";
        if (qs("adminClearPin")) qs("adminClearPin").checked = false;
        const tabsSel = qs("adminUserTabs");
        if (tabsSel) {
          if (Array.isArray(r.allowed_tabs) && r.allowed_tabs.length) {
            Array.from(tabsSel.options).forEach((o) => {
              o.selected = r.allowed_tabs.includes(o.value);
            });
          } else {
            Array.from(tabsSel.options).forEach((o) => {
              o.selected = false;
            });
          }
        }
      });
      tbody.appendChild(tr);
    });
    if (pre) pre.textContent = JSON.stringify({ count: rows.length }, null, 2);
    setStatus("Users loaded.");
  } catch (e) {
    if (pre) pre.textContent = String(e.message || e);
    setStatus("Failed to load users.");
  }
}

function escapeHtml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function applyAdminArtisanPreset() {
  const rolesSel = qs("adminRoles");
  const tabsSel = qs("adminUserTabs");
  if (rolesSel) {
    Array.from(rolesSel.options).forEach((o) => {
      o.selected = o.value === "artisan";
    });
  }
  const artisanTabs = getRoleAllowedTabs("artisan");
  if (tabsSel) {
    Array.from(tabsSel.options).forEach((o) => {
      o.selected = artisanTabs.includes(o.value);
    });
  }
  if (qs("adminDepartment") && !qs("adminDepartment").value) {
    qs("adminDepartment").value = "Workshop";
  }
  setStatus("Artisan preset applied — set a 4–6 digit PIN, then save user.");
}

async function saveAdminUser() {
  const username = String(qs("adminUsername")?.value || "").trim();
  const password = String(qs("adminPassword")?.value || "");
  const pin = String(qs("adminPin")?.value || "").replace(/\D/g, "");
  const clearPin = qs("adminClearPin")?.checked === true;
  const full_name = String(qs("adminFullName")?.value || "").trim();
  const department = String(qs("adminDepartment")?.value || "").trim();
  const allowedLocationsRaw = String(qs("adminAllowedLocations")?.value || "").trim();
  const issueSetup = qs("adminIssueSetupCode")?.checked !== false;
  const rolesSel = qs("adminRoles");
  const roles = rolesSel ? Array.from(rolesSel.selectedOptions).map((o) => String(o.value || "").trim().toLowerCase()).filter(Boolean) : [];
  if (!roles.length) return alert("Select at least one role.");
  const tabsSel = qs("adminUserTabs");
  const allowed_tabs = tabsSel ? Array.from(tabsSel.selectedOptions).map((o) => o.value) : [];
  if (!username) return alert("Username is required.");
  const body = {
    username,
    full_name: full_name || null,
    department: department || null,
    roles,
    role: roles[0],
    allowed_tabs,
    allowed_locations: allowedLocationsRaw || null,
    issue_setup_code: issueSetup,
  };
  if (password) body.password = password;
  if (clearPin) body.pin = "";
  else if (pin) {
    if (pin.length < 4 || pin.length > 6) return alert("PIN must be 4–6 digits.");
    body.pin = pin;
  }
  setStatus("Saving user…");
  try {
    const saved = await fetchJson(`${API}/api/auth/users`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const pre = qs("adminUsersResult");
    if (saved?.setup_code) {
      if (pre) pre.textContent = `Setup code for ${username}: ${saved.setup_code}\nExpires: ${saved.setup_code_expires_at || "7 days"}`;
      setStatus("User saved. Share setup code with user.");
    } else {
      setStatus("User saved.");
    }
    await loadAdminUsers();
  } catch (e) {
    setStatus("Save user failed: " + (e.message || e));
  }
}

let safetyTplItems = [];
let safetyReportSelectedCodes = new Set();
let lastSafetyQrUrl = "";

function setSafetyAdminResult(text) {
  const pre = qs("safetyAdminResult");
  if (pre) pre.textContent = String(text || "");
}

function renderSafetyTemplateEditor(items) {
  safetyTplItems = Array.isArray(items) ? items.map((it) => ({
    key: String(it.key || ""),
    label: String(it.label || ""),
  })) : [];
  const host = qs("safetyTplItems");
  if (!host) return;
  if (!safetyTplItems.length) {
    host.innerHTML = `<div class="muted small">No checklist rows.</div>`;
    return;
  }
  host.innerHTML = safetyTplItems.map((it, idx) => `
    <div class="item" style="display:flex; gap:8px; flex-wrap:wrap; align-items:center; margin-bottom:6px;">
      <input type="text" class="w-140" data-safety-tpl-key="${idx}" value="${escapeHtml(it.key)}" placeholder="key" />
      <input type="text" style="flex:1; min-width:200px;" data-safety-tpl-label="${idx}" value="${escapeHtml(it.label)}" placeholder="Label" />
      <button type="button" class="btn btn-secondary btn-sm" data-safety-tpl-remove="${idx}">Remove</button>
    </div>
  `).join("");
}

async function loadSafetyTemplatesSelect(preferredKey) {
  const data = await fetchJson(`${API}/api/safety/templates`);
  const templates = Array.isArray(data.templates) ? data.templates : [];
  const selects = ["safetyTplSelect", "safetyItemTemplate", "safetyPdfType"].map((id) => qs(id)).filter(Boolean);
  const prevKey = String(preferredKey || qs("safetyTplSelect")?.value || "").trim();
  selects.forEach((sel) => {
    const keepAll = sel.id === "safetyPdfType";
    const opts = templates.map((t) =>
      `<option value="${escapeHtml(t.template_key)}">${escapeHtml(t.title || t.template_key)}</option>`
    );
    sel.innerHTML = keepAll
      ? `<option value="">All types</option>${opts.join("")}`
      : opts.join("");
    if (prevKey && templates.some((t) => t.template_key === prevKey)) {
      sel.value = prevKey;
    }
  });
  renderSafetyCategoriesList(templates);
  return templates;
}

function renderSafetyCategoriesList(templates) {
  const host = qs("safetyCategoriesList");
  if (!host) return;
  const rows = Array.isArray(templates) ? templates : [];
  if (!rows.length) {
    host.innerHTML = `<div class="muted small">No categories yet — add one above.</div>`;
    return;
  }
  host.innerHTML = rows
    .map(
      (t) => `
    <div class="item safety-category-row" style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; align-items:center;">
      <div>
        <strong>${escapeHtml(t.title || t.template_key)}</strong>
        <div class="muted small"><code>${escapeHtml(t.template_key)}</code> · ${Number(t.item_count || 0)} item(s) · ${Number(t.items?.length || 0)} checklist row(s)</div>
      </div>
      <button type="button" class="btn btn-secondary btn-sm" data-safety-edit-category="${escapeHtml(t.template_key)}">Edit checklist</button>
    </div>`
    )
    .join("");
}

async function addSafetyCategory() {
  const title = String(qs("safetyCategoryTitle")?.value || "").trim();
  const template_key = String(qs("safetyCategoryKey")?.value || "").trim();
  if (!title) return alert("Enter a category name (e.g. Cutting equipment).");
  setStatus("Adding safety category…");
  const body = { title };
  if (template_key) body.template_key = template_key;
  const data = await fetchJson(`${API}/api/safety/templates`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const key = String(data?.template?.template_key || template_key || "").trim();
  if (qs("safetyCategoryTitle")) qs("safetyCategoryTitle").value = "";
  if (qs("safetyCategoryKey")) qs("safetyCategoryKey").value = "";
  await loadSafetyTemplatesSelect(key);
  if (qs("safetyTplSelect") && key) qs("safetyTplSelect").value = key;
  if (qs("safetyItemTemplate") && key) qs("safetyItemTemplate").value = key;
  await loadSafetyTemplateEditor();
  setSafetyAdminResult(`Added category "${title}" (${key || "saved"}). Edit checklist rows below.`);
  setStatus("Safety category added.");
}

async function loadSafetyTemplateEditor() {
  const key = String(qs("safetyTplSelect")?.value || "fire_extinguisher").trim();
  setStatus("Loading safety template…");
  const data = await fetchJson(`${API}/api/safety/templates/${encodeURIComponent(key)}`);
  renderSafetyTemplateEditor(data?.template?.items || []);
  setSafetyAdminResult(`Loaded template: ${key}`);
  setStatus("Safety template loaded.");
}

async function saveSafetyTemplateEditor() {
  const key = String(qs("safetyTplSelect")?.value || "fire_extinguisher").trim();
  const items = safetyTplItems.map((_, idx) => {
    const keyInp = document.querySelector(`input[data-safety-tpl-key="${idx}"]`);
    const labelInp = document.querySelector(`input[data-safety-tpl-label="${idx}"]`);
    return {
      key: String(keyInp?.value || "").trim(),
      label: String(labelInp?.value || "").trim(),
    };
  }).filter((r) => r.key && r.label);
  if (!items.length) return alert("Add at least one checklist row.");
  setStatus("Saving safety template…");
  const data = await fetchJson(`${API}/api/safety/templates/${encodeURIComponent(key)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ items }),
  });
  renderSafetyTemplateEditor(data?.template?.items || items);
  await loadSafetyTemplatesSelect(key);
  setSafetyAdminResult(`Saved template ${key} (${items.length} rows).`);
  setStatus("Safety template saved.");
}

async function loadSafetyItemsList() {
  const host = qs("safetyItemsList");
  if (!host) return;
  const data = await fetchJson(`${API}/api/safety/items`);
  const rows = Array.isArray(data.items) ? data.items : [];
  if (!rows.length) {
    safetyReportSelectedCodes = new Set();
    host.innerHTML = `<div class="muted small">No safety items registered yet.</div>`;
    return;
  }
  const validCodes = new Set(rows.map((r) => String(r.item_code || "").trim().toUpperCase()).filter(Boolean));
  safetyReportSelectedCodes = new Set([...safetyReportSelectedCodes].filter((c) => validCodes.has(c)));
  host.innerHTML = rows.map((r) => `
    <div class="item" style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; align-items:center;">
      <div>
        <label style="display:flex; align-items:flex-start; gap:8px;">
          <input type="checkbox" data-safety-report-select="${escapeHtml(r.item_code)}" ${safetyReportSelectedCodes.has(String(r.item_code || "").trim().toUpperCase()) ? "checked" : ""} />
          <span>
            <strong>${escapeHtml(r.item_code)}</strong> — ${escapeHtml(r.item_name || r.template_title || "")}
            <div class="muted small">
              ${escapeHtml(r.template_title || r.template_key || "")}${r.location ? ` · ${escapeHtml(r.location)}` : ""}
              ${String(r.latest_status || "").toLowerCase() === "fail" ? ` · <span class="pill pill-red">FLAGGED</span>` : ""}
            </div>
          </span>
        </label>
      </div>
      <div class="row stack-10">
        <button type="button" class="btn btn-secondary btn-sm" data-safety-item-pdf="${escapeHtml(r.item_code)}" title="Individual inspection PDF">PDF</button>
        <button type="button" class="btn btn-secondary btn-sm" data-safety-use-qr="${escapeHtml(r.item_code)}">QR</button>
        <button type="button" class="btn btn-secondary btn-sm" data-safety-open-insp="${escapeHtml(r.item_code)}">Inspect</button>
        <button type="button" class="btn btn-secondary btn-sm" data-safety-remove-item="${Number(r.id)}">Remove</button>
      </div>
    </div>
  `).join("");
}

async function addSafetyEquipmentItem() {
  const item_code = String(qs("safetyItemCode")?.value || "").trim();
  const template_key = String(qs("safetyItemTemplate")?.value || "fire_extinguisher").trim();
  const item_name = String(qs("safetyItemName")?.value || "").trim();
  const location = String(qs("safetyItemLocation")?.value || "").trim();
  if (!item_code) return alert("Item code is required (e.g. FE-WS-01).");
  setStatus("Adding safety item…");
  await fetchJson(`${API}/api/safety/items`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ item_code, template_key, item_name, location }),
  });
  if (qs("safetyItemCode")) qs("safetyItemCode").value = "";
  if (qs("safetyItemName")) qs("safetyItemName").value = "";
  if (qs("safetyItemLocation")) qs("safetyItemLocation").value = "";
  await loadSafetyItemsList();
  await loadSafetyTemplatesSelect(template_key);
  setSafetyAdminResult(`Added ${item_code.toUpperCase()}.`);
  setStatus("Safety item added.");
}

async function removeSafetyEquipmentItem(id) {
  const rowId = Number(id || 0);
  if (!rowId) return;
  if (!window.confirm("Remove this safety item from the register?")) return;
  setStatus("Removing safety item…");
  await fetchJson(`${API}/api/safety/items/${rowId}`, { method: "DELETE" });
  await loadSafetyItemsList();
  await loadSafetyTemplatesSelect();
  setStatus("Safety item removed.");
}

async function buildSafetyQrImageData(itemCode) {
  const code = String(itemCode || "").trim().toUpperCase();
  if (!code) throw new Error("Item code is required.");
  const res = await fetchJson(`${API}/api/safety/items/${encodeURIComponent(code)}/qr-profile/refresh`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({}),
  });
  const scanValue = String(res?.qr_payload?.scan_url || "").trim();
  if (!scanValue) throw new Error("No QR scan URL generated.");
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=420x420&data=${encodeURIComponent(scanValue)}`;
  return { qrUrl, qrText: String(res?.qr_text || ""), scanValue };
}

async function generateSafetyQr() {
  const code = String(qs("safetyQrItemCode")?.value || qs("safetyItemCode")?.value || "").trim();
  if (!code) return alert("Enter a safety item code.");
  setStatus(`Generating QR for ${code}…`);
  const { qrUrl, qrText } = await buildSafetyQrImageData(code);
  lastSafetyQrUrl = qrUrl;
  const prev = qs("safetyQrPreview");
  const img = qs("safetyQrImg");
  const txt = qs("safetyQrText");
  if (img) img.src = qrUrl;
  if (txt) txt.textContent = qrText;
  if (prev) prev.style.display = "block";
  setSafetyAdminResult(`QR ready for ${code.toUpperCase()}.`);
  setStatus(`QR generated for ${code.toUpperCase()}.`);
}

function printSafetyQr() {
  if (!lastSafetyQrUrl) {
    alert("Generate a QR first.");
    return;
  }
  const code = String(qs("safetyQrItemCode")?.value || "").trim().toUpperCase();
  openQrLabelSheetPrintWindow(
    [{ code: code || "Safety item", qrUrl: lastSafetyQrUrl }],
    readQrSheetLayout("safety"),
    "IRONLOG Safety QR Label Sheet"
  );
  setStatus("Safety QR print sheet opened.");
}

async function printAllSafetyQrSheet() {
  setStatus("Building safety QR label sheet…");
  const data = await fetchJson(`${API}/api/safety/items`);
  const rows = Array.isArray(data.items) ? data.items : [];
  if (!rows.length) return alert("No safety items registered yet.");
  const labels = [];
  for (let i = 0; i < rows.length; i += 1) {
    const code = String(rows[i]?.item_code || "").trim();
    if (!code) continue;
    try {
      const { qrUrl } = await buildSafetyQrImageData(code);
      labels.push({ code, qrUrl });
      setStatus(`Preparing safety label ${i + 1}/${rows.length}: ${code}`);
      await new Promise((resolve) => setTimeout(resolve, 80));
    } catch {
      /* skip failed rows */
    }
  }
  if (!labels.length) throw new Error("Could not prepare any safety QR labels.");
  openQrLabelSheetPrintWindow(labels, readQrSheetLayout("safety"), "IRONLOG Safety QR Label Sheet");
  setStatus(`Safety QR sheet ready (${labels.length} labels) ✅`);
}

async function openSafetyRegisterPdf(blank = false) {
  const template_key = String(qs("safetyPdfType")?.value || "").trim();
  const date = String(qs("safetyPdfDate")?.value || "").trim() || new Date().toISOString().slice(0, 10);
  const q = new URLSearchParams();
  if (template_key) q.set("template_key", template_key);
  q.set("date", date);
  if (blank) q.set("blank", "1");
  setStatus(blank ? "Opening blank safety sheet…" : "Opening safety register PDF…");
  await openAuthedPdf(`${API}/api/safety/register.pdf?${q.toString()}`);
  setStatus("Safety PDF opened.");
}

async function openSafetyItemInspectionPdf(itemCode) {
  const code = String(itemCode || "").trim().toUpperCase();
  if (!code) return;
  const date =
    String(qs("safetyReportEndDate")?.value || "").trim() ||
    String(qs("safetyPdfDate")?.value || "").trim() ||
    new Date().toISOString().slice(0, 10);
  const q = new URLSearchParams();
  q.set("item_code", code);
  q.set("date", date);
  setStatus(`Opening safety PDF for ${code}…`);
  await openAuthedPdf(`${API}/api/safety/inspections/item.pdf?${q.toString()}`);
  setStatus(`Safety PDF opened for ${code}.`);
}

async function openSafetyInspectionReportPdf(selectedOnly = false) {
  const start = String(qs("safetyReportStartDate")?.value || "").trim() || new Date().toISOString().slice(0, 10);
  const end = String(qs("safetyReportEndDate")?.value || "").trim() || start;
  const template_key = String(qs("safetyPdfType")?.value || "").trim();
  const selectedCodes = [...safetyReportSelectedCodes];
  if (selectedOnly && !selectedCodes.length) {
    alert("Select at least one equipment item in the register list first.");
    return;
  }
  const q = new URLSearchParams();
  q.set("start", start);
  q.set("end", end);
  if (template_key) q.set("template_key", template_key);
  if (selectedOnly && selectedCodes.length) q.set("item_codes", selectedCodes.join(","));
  setStatus(selectedOnly ? "Opening selected safety inspection report PDF…" : "Opening safety inspection report PDF…");
  await openAuthedPdf(`${API}/api/safety/inspections/report.pdf?${q.toString()}`);
  setStatus("Safety inspection report opened.");
}

async function initSafetyAdminPanel() {
  if (!qs("adminSafetyCard")) return;
  const pdfDate = qs("safetyPdfDate");
  if (pdfDate && !pdfDate.value) pdfDate.value = new Date().toISOString().slice(0, 10);
  const reportStart = qs("safetyReportStartDate");
  const reportEnd = qs("safetyReportEndDate");
  if (reportEnd && !reportEnd.value) reportEnd.value = new Date().toISOString().slice(0, 10);
  if (reportStart && !reportStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    reportStart.value = d.toISOString().slice(0, 10);
  }
  applySafetyQrSheetPreset();
  try {
  await loadSafetyTemplatesSelect();
  await loadSafetyTemplateEditor();
    await loadSafetyItemsList();
  } catch (e) {
    setSafetyAdminResult(String(e.message || e));
  }
}

function setTelemAdminResult(text) {
  const pre = qs("telemAdminResult");
  if (pre) pre.textContent = String(text || "");
}

function telemLinkStatusPill(status) {
  const st = String(status || "offline").toLowerCase();
  if (st === "live") return `<span class="pill green" style="font-size:0.65rem;">LIVE</span>`;
  if (st === "stale") return `<span class="pill amber" style="font-size:0.65rem;">STALE</span>`;
  if (st === "inactive") return `<span class="pill" style="font-size:0.65rem;">INACTIVE</span>`;
  return `<span class="pill" style="font-size:0.65rem;">OFFLINE</span>`;
}

async function loadTelematicsAdminDevices() {
  const host = qs("telemDevicesList");
  if (!host) return;
  const showInactive = qs("telemShowInactive")?.checked === true;
  setStatus("Loading telematics units…");
  try {
    const q = showInactive ? "?all=1" : "";
    const data = await fetchJson(`${API}/api/telematics/devices${q}`);
    const rows = Array.isArray(data.devices) ? data.devices : [];
    if (!rows.length) {
      host.innerHTML = `<div class="muted small">No telematics units registered yet. Add one above.</div>`;
      setStatus("No telematics units.");
      return;
    }
    host.innerHTML = rows.map((r) => {
      const active = Number(r.active) === 1 || r.active === true;
      const meter = r.engine_hours == null ? "—" : `${Number(r.engine_hours).toFixed(1)} h`;
      const lastSeen = String(r.recorded_at || r.snapshot_updated_at || r.updated_at || "—");
      return `
        <div class="item" style="display:flex; justify-content:space-between; gap:10px; flex-wrap:wrap; align-items:center; opacity:${active ? "1" : "0.65"};">
          <div>
            <strong>${escapeHtml(r.asset_code)}</strong> ${telemLinkStatusPill(active ? r.link_status : "inactive")}
            <div class="muted small">${escapeHtml(r.asset_name || "")} · ${escapeHtml(r.unit_model || "FSC")} · S/N ${escapeHtml(r.device_serial || "")}</div>
            <div class="muted small">Meter: ${escapeHtml(meter)} · Last seen: ${escapeHtml(lastSeen)}</div>
          </div>
          <div class="row stack-10">
            <button type="button" class="btn btn-secondary btn-sm" data-telem-edit="${Number(r.id)}" data-telem-asset="${escapeHtml(r.asset_code)}" data-telem-serial="${escapeHtml(r.device_serial || "")}" data-telem-model="${escapeHtml(r.unit_model || "FSC650")}" data-telem-ext="${escapeHtml(r.external_id || "")}">Replace unit</button>
            ${active ? `<button type="button" class="btn btn-secondary btn-sm" data-telem-deactivate="${Number(r.id)}" data-telem-asset-label="${escapeHtml(r.asset_code)}">Deactivate</button>` : ""}
          </div>
        </div>
      `;
    }).join("");
    setStatus(`Telematics units loaded (${rows.length}).`);
  } catch (e) {
    host.innerHTML = `<div class="muted small">Load failed: ${escapeHtml(e.message || String(e))}</div>`;
    setStatus("Telematics load failed.");
  }
}

function fillTelematicsDeviceForm({ assetCode, deviceSerial, unitModel, externalId, replaceFaulty = false }) {
  if (qs("telemAssetCode")) qs("telemAssetCode").value = String(assetCode || "");
  if (qs("telemDeviceSerial")) qs("telemDeviceSerial").value = String(deviceSerial || "");
  if (qs("telemUnitModel")) qs("telemUnitModel").value = String(unitModel || "FSC650");
  if (qs("telemExternalId")) qs("telemExternalId").value = String(externalId || "");
  if (qs("telemReplaceFaulty")) qs("telemReplaceFaulty").checked = Boolean(replaceFaulty);
  qs("telemDeviceSerial")?.focus();
}

async function saveTelematicsDevice() {
  const asset_code = String(qs("telemAssetCode")?.value || "").trim();
  const device_serial = String(qs("telemDeviceSerial")?.value || "").trim();
  const unit_model = String(qs("telemUnitModel")?.value || "FSC650").trim();
  const external_id = String(qs("telemExternalId")?.value || "").trim();
  const replace_faulty = qs("telemReplaceFaulty")?.checked === true;
  if (!asset_code || !device_serial) return alert("Asset code and device serial are required.");
  setStatus("Saving telematics unit…");
  const data = await fetchJson(`${API}/api/telematics/devices`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      asset_code,
      device_serial,
      unit_model,
      external_id: external_id || device_serial,
      replace_faulty,
    }),
  });
  const msg = data.replaced
    ? `Replaced unit on ${asset_code.toUpperCase()} → serial ${device_serial}.`
    : data.created
      ? `Registered ${device_serial} on ${asset_code.toUpperCase()}.`
      : `Updated ${asset_code.toUpperCase()}.`;
  setTelemAdminResult(msg);
  if (qs("telemReplaceFaulty")) qs("telemReplaceFaulty").checked = false;
  await loadTelematicsAdminDevices();
  loadTelematicsFleet().catch(() => {});
  setStatus(msg);
}

async function deactivateTelematicsDeviceAdmin(id, assetLabel) {
  const deviceId = Number(id || 0);
  if (!deviceId) return;
  const label = String(assetLabel || "this asset").trim();
  if (!window.confirm(`Deactivate telematics unit for ${label}? Daily hours will unlock until a new unit is registered and reports.`)) return;
  setStatus("Deactivating telematics unit…");
  const data = await fetchJson(`${API}/api/telematics/devices/${deviceId}/deactivate`, {
    method: "POST",
    headers: authHeaders(),
  });
  setTelemAdminResult(`Deactivated ${data.device_serial || "unit"} on ${data.asset_code || label}.`);
  await loadTelematicsAdminDevices();
  loadTelematicsFleet().catch(() => {});
  setStatus("Telematics unit deactivated.");
}

async function initTelematicsAdminPanel() {
  if (!qs("adminTelematicsCard")) return;
  await loadTelematicsAdminDevices().catch((e) => setTelemAdminResult(String(e.message || e)));
}

function applyMdmPolicyCheckboxes(policies) {
  const p = policies && typeof policies === "object" ? policies : {};
  const asset = Array.isArray(p.asset) ? p.asset : [];
  const part = Array.isArray(p.part_stock_intake) ? p.part_stock_intake : [];
  const setChk = (id, key, list) => {
    const el = qs(id);
    if (el) el.checked = list.includes(key);
  };
  setChk("mdmPolAssetDept", "department_code", asset);
  setChk("mdmPolAssetCc", "cost_center_code", asset);
  setChk("mdmPolAssetOwner", "data_owner_username", asset);
  setChk("mdmPolPartDept", "department_code", part);
  setChk("mdmPolPartSup", "default_supplier_code", part);
  setChk("mdmPolPartOwner", "data_owner_username", part);
}

async function saveMdmPolicies() {
  setStatus("Saving policies…");
  try {
    const policies = {
      asset: [],
      part_stock_intake: [],
    };
    if (qs("mdmPolAssetDept")?.checked) policies.asset.push("department_code");
    if (qs("mdmPolAssetCc")?.checked) policies.asset.push("cost_center_code");
    if (qs("mdmPolAssetOwner")?.checked) policies.asset.push("data_owner_username");
    if (qs("mdmPolPartDept")?.checked) policies.part_stock_intake.push("department_code");
    if (qs("mdmPolPartSup")?.checked) policies.part_stock_intake.push("default_supplier_code");
    if (qs("mdmPolPartOwner")?.checked) policies.part_stock_intake.push("data_owner_username");
    const res = await fetchJson(`${API}/api/masterdata/policies`, {
      method: "PUT",
      body: JSON.stringify({ policies }),
    });
    applyMdmPolicyCheckboxes(res.policies);
    setStatus("Policies saved.");
  } catch (e) {
    setStatus("Policy save failed: " + (e.message || e));
  }
}

async function loadMasterDataGovernance() {
  const sumEl = qs("mdmSummaryOut");
  const dBody = qs("mdmDeptTbody");
  const cBody = qs("mdmCcTbody");
  const sBody = qs("mdmSupTbody");
  if (!dBody || !cBody || !sBody) return;
  setStatus("Loading master data…");
  try {
    const [sum, dep, cc, sup, pol] = await Promise.all([
      fetchJson(`${API}/api/masterdata/summary`),
      fetchJson(`${API}/api/masterdata/departments`),
      fetchJson(`${API}/api/masterdata/cost-centers`),
      fetchJson(`${API}/api/masterdata/suppliers`),
      fetchJson(`${API}/api/masterdata/policies`),
    ]);
    applyMdmPolicyCheckboxes(pol.policies);
    if (sumEl) {
      sumEl.textContent = JSON.stringify(
        { site: sum?.site_code, counts: sum?.counts },
        null,
        2
      );
    }
    const dRows = Array.isArray(dep.rows) ? dep.rows : [];
    dBody.innerHTML = dRows
      .map(
        (r) =>
          `<tr><td>${escapeHtml(r.code)}</td><td>${escapeHtml(r.name || "")}</td><td>${escapeHtml(r.owner_username || "—")}</td><td>${r.active ? "yes" : "no"}</td></tr>`
      )
      .join("");
    const cRows = Array.isArray(cc.rows) ? cc.rows : [];
    cBody.innerHTML = cRows
      .map(
        (r) =>
          `<tr><td>${escapeHtml(r.code)}</td><td>${escapeHtml(r.name || "")}</td><td>${escapeHtml(r.department_code || "—")}</td><td>${r.active ? "yes" : "no"}</td></tr>`
      )
      .join("");
    const sRows = Array.isArray(sup.rows) ? sup.rows : [];
    sBody.innerHTML = sRows
      .map(
        (r) =>
          `<tr><td>${escapeHtml(r.supplier_code)}</td><td>${escapeHtml(r.name || "")}</td><td>${escapeHtml(r.contact_email || "—")}</td><td>${r.active ? "yes" : "no"}</td></tr>`
      )
      .join("");
    setStatus("Master data loaded.");
  } catch (e) {
    if (sumEl) sumEl.textContent = String(e.message || e);
    setStatus("Master data load failed.");
  }
}

async function saveMdmDepartment() {
  const code = String(qs("mdmDeptCode")?.value || "").trim();
  const name = String(qs("mdmDeptName")?.value || "").trim();
  const owner_username = String(qs("mdmDeptOwner")?.value || "").trim() || null;
  if (!code || !name) return alert("Department code and name are required.");
  setStatus("Saving department…");
  try {
    await fetchJson(`${API}/api/masterdata/departments`, {
      method: "POST",
      body: JSON.stringify({ code, name, owner_username }),
    });
    if (qs("mdmDeptCode")) qs("mdmDeptCode").value = "";
    if (qs("mdmDeptName")) qs("mdmDeptName").value = "";
    if (qs("mdmDeptOwner")) qs("mdmDeptOwner").value = "";
    await loadMasterDataGovernance();
    setStatus("Department saved.");
  } catch (e) {
    setStatus("Department save failed: " + (e.message || e));
  }
}

async function saveMdmCostCenter() {
  const code = String(qs("mdmCcCode")?.value || "").trim();
  const name = String(qs("mdmCcName")?.value || "").trim();
  const department_code = String(qs("mdmCcDept")?.value || "").trim() || null;
  if (!code || !name) return alert("Cost center code and name are required.");
  setStatus("Saving cost center…");
  try {
    await fetchJson(`${API}/api/masterdata/cost-centers`, {
      method: "POST",
      body: JSON.stringify({ code, name, department_code }),
    });
    if (qs("mdmCcCode")) qs("mdmCcCode").value = "";
    if (qs("mdmCcName")) qs("mdmCcName").value = "";
    if (qs("mdmCcDept")) qs("mdmCcDept").value = "";
    await loadMasterDataGovernance();
    setStatus("Cost center saved.");
  } catch (e) {
    setStatus("Cost center save failed: " + (e.message || e));
  }
}

async function saveMdmSupplier() {
  const supplier_code = String(qs("mdmSupCode")?.value || "").trim();
  const name = String(qs("mdmSupName")?.value || "").trim();
  const contact_email = String(qs("mdmSupEmail")?.value || "").trim() || null;
  if (!supplier_code || !name) return alert("Supplier code and name are required.");
  setStatus("Saving supplier…");
  try {
    await fetchJson(`${API}/api/masterdata/suppliers`, {
      method: "POST",
      body: JSON.stringify({ supplier_code, name, contact_email }),
    });
    if (qs("mdmSupCode")) qs("mdmSupCode").value = "";
    if (qs("mdmSupName")) qs("mdmSupName").value = "";
    if (qs("mdmSupEmail")) qs("mdmSupEmail").value = "";
    await loadMasterDataGovernance();
    setStatus("Supplier saved.");
  } catch (e) {
    setStatus("Supplier save failed: " + (e.message || e));
  }
}

async function submitChangePassword() {
  const old_password = String(qs("chPwdOld")?.value || "");
  const new_password = String(qs("chPwdNew")?.value || "").trim();
  const out = qs("chPwdResult");
  if (out) out.textContent = "";
  if (new_password.length < 6) return alert("New password must be at least 6 characters.");
  setStatus("Updating password…");
  try {
    await fetchJson(`${API}/api/auth/change-password`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ old_password, new_password }),
    });
    if (qs("chPwdOld")) qs("chPwdOld").value = "";
    if (qs("chPwdNew")) qs("chPwdNew").value = "";
    if (out) out.textContent = "Password updated.";
    setStatus("Password updated.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Password change failed.");
  }
}

let smtpHasStoredPassword = false;

async function loadSmtpSettings() {
  const out = qs("smtpSettingsResult");
  try {
    const data = await fetchJson(`${API}/api/reports/smtp-settings`);
    const s = data?.settings || {};
    smtpHasStoredPassword = Boolean(s.has_password);
    if (qs("smtpHost")) qs("smtpHost").value = String(s.host || "");
    if (qs("smtpPort")) qs("smtpPort").value = String(Number(s.port || 587));
    if (qs("smtpSecure")) qs("smtpSecure").value = Number(s.secure || 0) === 1 ? "1" : "0";
    if (qs("smtpUsername")) qs("smtpUsername").value = String(s.username || "");
    if (qs("smtpPassword")) qs("smtpPassword").value = "";
    if (qs("smtpFromEmail")) qs("smtpFromEmail").value = String(s.from_email || "");
    if (qs("smtpFromName")) qs("smtpFromName").value = String(s.from_name || "");
    if (out) {
      out.textContent = `Loaded SMTP settings.\nPassword set: ${s.has_password ? "yes" : "no"}\nUpdated by: ${s.updated_by || "-"}\nUpdated at: ${s.updated_at || "-"}`;
    }
    await loadSmtpSubscriptionOptions();
    setStatus("SMTP settings loaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("SMTP settings load failed.");
  }
}

async function loadSmtpSubscriptionOptions() {
  const sel = qs("smtpSubscriptionId");
  if (!sel) return;
  const current = String(sel.value || "");
  const data = await fetchJson(`${API}/api/reports/subscriptions`);
  const rows = Array.isArray(data?.subscriptions) ? data.subscriptions : [];
  sel.innerHTML = '<option value="">Select subscription...</option>';
  rows.forEach((r) => {
    const id = Number(r.id || 0);
    if (!id) return;
    const opt = document.createElement("option");
    opt.value = String(id);
    const channel = String(r.channel || "").toLowerCase();
    const active = Number(r.active || 0) === 1 ? "active" : "paused";
    opt.textContent = `${r.name || "Subscription"} (#${id}) - ${channel} - ${active}`;
    sel.appendChild(opt);
  });
  if (current && Array.from(sel.options).some((o) => o.value === current)) sel.value = current;
}

async function saveSmtpSettings() {
  const out = qs("smtpSettingsResult");
  const body = {
    host: String(qs("smtpHost")?.value || "").trim(),
    port: Number(qs("smtpPort")?.value || 587),
    secure: Number(qs("smtpSecure")?.value || 0) === 1 ? 1 : 0,
    username: String(qs("smtpUsername")?.value || "").trim(),
    password: String(qs("smtpPassword")?.value || ""),
    from_email: String(qs("smtpFromEmail")?.value || "").trim(),
    from_name: String(qs("smtpFromName")?.value || "").trim(),
  };
  if (!body.host) return alert("SMTP host is required.");
  if (!body.username) return alert("SMTP username is required.");
  if (!body.from_email) return alert("From email is required.");
  if (!body.password && !smtpHasStoredPassword) {
    return alert("SMTP password is required on first setup. Enter the password, then Save SMTP.");
  }
  setStatus("Saving SMTP settings…");
  try {
    const data = await fetchJson(`${API}/api/reports/smtp-settings`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (qs("smtpPassword")) qs("smtpPassword").value = "";
    smtpHasStoredPassword = Boolean(data?.settings?.has_password);
    if (out) out.textContent = JSON.stringify({ ok: true, settings: data?.settings || {} }, null, 2);
    setStatus("SMTP settings saved.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("SMTP save failed.");
  }
}

async function testSmtpSettings() {
  const out = qs("smtpSettingsResult");
  const to = String(qs("smtpTestTo")?.value || "").trim();
  if (!to) return alert("Enter a test recipient email.");
  setStatus("Sending SMTP test email…");
  try {
    const data = await fetchJson(`${API}/api/reports/smtp-settings/test`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ to }),
    });
    if (out) out.textContent = JSON.stringify(data, null, 2);
    setStatus("SMTP test email sent.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("SMTP test failed.");
  }
}

function renderPushNotifyDevices(devices) {
  const tbody = qs("pushNotifyDevicesTbody");
  if (!tbody) return;
  const rows = Array.isArray(devices) ? devices : [];
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="4" class="muted small">No devices registered yet. Technicians must sign into IRONLOG Notify on their phone.</td></tr>';
    return;
  }
  tbody.innerHTML = rows
    .map((d) => {
      const user = escapeHtml(String(d.username || ""));
      const label = escapeHtml(String(d.device_label || "—"));
      const platform = escapeHtml(String(d.platform || "android"));
      const seen = escapeHtml(String(d.last_seen_at || "—"));
      return `<tr><td>${user}</td><td>${label}</td><td>${platform}</td><td>${seen}</td></tr>`;
    })
    .join("");
}

function renderPushNotifyRecent(recent) {
  const tbody = qs("pushNotifyRecentTbody");
  if (!tbody) return;
  const rows = Array.isArray(recent) ? recent : [];
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="muted small">No notifications sent yet.</td></tr>';
    return;
  }
  tbody.innerHTML = rows
    .map((r) => {
      const when = escapeHtml(String(r.sent_at || "—"));
      const user = escapeHtml(String(r.username || "—"));
      const kindVal = escapeHtml(String(r.kind || "—"));
      const title = escapeHtml(String(r.title || "—"));
      const ok = r.success ? "Yes" : "No";
      const err = r.error ? ` title="${escapeHtml(String(r.error))}"` : "";
      return `<tr><td>${when}</td><td>${user}</td><td>${kindVal}</td><td>${title}</td><td${err}>${ok}</td></tr>`;
    })
    .join("");
}

function populatePushNotifyUserOptions(devices) {
  const sel = qs("pushNotifyUsername");
  if (!sel) return;
  const current = String(sel.value || "");
  const rows = Array.isArray(devices) ? devices : [];
  const seen = new Set();
  sel.innerHTML = '<option value="">Select technician…</option>';
  rows.forEach((d) => {
    const username = String(d.username || "").trim();
    if (!username || seen.has(username.toLowerCase())) return;
    seen.add(username.toLowerCase());
    const opt = document.createElement("option");
    opt.value = username;
    const label = String(d.device_label || "").trim();
    opt.textContent = label ? `${username} (${label})` : username;
    sel.appendChild(opt);
  });
  if (current && Array.from(sel.options).some((o) => o.value === current)) sel.value = current;
}

async function loadPushNotificationSettings() {
  const out = qs("pushNotifyResult");
  const statusEl = qs("pushNotifyStatus");
  try {
    const data = await fetchJson(`${API}/api/notifications/admin`);
    const enabled = Boolean(data.push_enabled);
    if (statusEl) {
      statusEl.textContent = enabled
        ? "Server push is configured and ready."
        : "Push is not configured on the server (Firebase service account missing). Devices can register but alerts will not send.";
      statusEl.className = enabled ? "muted small" : "muted small";
      statusEl.style.color = enabled ? "" : "#b45309";
    }
    const apkLink = qs("pushNotifyApkLink");
    if (apkLink && data.apk_url) apkLink.href = data.apk_url;
    const expoLink = qs("pushNotifyExpoLink");
    if (expoLink && data.expo_install_url) {
      expoLink.href = data.expo_install_url;
      expoLink.textContent = "Open Expo install page";
    }
    populatePushNotifyUserOptions(data.devices);
    renderPushNotifyDevices(data.devices);
    renderPushNotifyRecent(data.recent);
    if (out) {
      out.textContent =
        `Push enabled: ${enabled ? "yes" : "no"}\n` +
        `Registered devices: ${Array.isArray(data.devices) ? data.devices.length : 0}\n` +
        `APK: ${data.apk_url || "(not set)"}`;
    }
    setStatus("Push notification settings loaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    if (statusEl) statusEl.textContent = "Could not load push notification settings.";
    setStatus("Push notification load failed.");
  }
}

async function sendPushNotificationTest() {
  const out = qs("pushNotifyResult");
  const username = String(qs("pushNotifyUsername")?.value || "").trim();
  if (!username) return alert("Select a technician with a registered device.");
  setStatus("Sending push test…");
  try {
    const data = await fetchJson(`${API}/api/notifications/test`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username }),
    });
    if (out) out.textContent = JSON.stringify(data, null, 2);
    setStatus(data.sent ? "Push test sent." : "Push test completed (check result — device may be unregistered).");
    await loadPushNotificationSettings();
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Push test failed.");
  }
}

async function sendPushNotificationManual() {
  const out = qs("pushNotifyResult");
  const username = String(qs("pushNotifyUsername")?.value || "").trim();
  const title = String(qs("pushNotifyTitle")?.value || "").trim();
  const body = String(qs("pushNotifyBody")?.value || "").trim();
  const woId = String(qs("pushNotifyWoId")?.value || "").trim();
  if (!username) return alert("Select a technician with a registered device.");
  if (!title) return alert("Enter a notification title.");
  if (!body) return alert("Enter a notification message.");
  setStatus("Sending push notification…");
  try {
    const payload = { username, title, body };
    if (woId) payload.wo_id = woId;
    const data = await fetchJson(`${API}/api/notifications/send`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (out) out.textContent = JSON.stringify(data, null, 2);
    setStatus(data.sent ? "Push notification sent." : "Send completed (check result).");
    await loadPushNotificationSettings();
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Push send failed.");
  }
}

let pdfReportSitesCache = [];

function updatePdfReportLogoPreview(available, customOnly = false) {
  const img = qs("pdfReportLogoPreview");
  const removeBtn = qs("removePdfReportLogoBtn");
  if (!img) return;
  if (available) {
    const headers = new Headers(authHeaders());
    const tok = getAuthToken();
    if (tok) headers.set("Authorization", `Bearer ${tok}`);
    fetch(`${API}/api/reports/pdf-settings/logo?t=${Date.now()}`, { headers })
      .then((res) => {
        if (!res.ok) throw new Error("Logo not found");
        return res.blob();
      })
      .then((blob) => {
        const url = URL.createObjectURL(blob);
        if (img.dataset.blobUrl) URL.revokeObjectURL(img.dataset.blobUrl);
        img.dataset.blobUrl = url;
        img.src = url;
        img.hidden = false;
      })
      .catch(() => {
        img.hidden = true;
        img.removeAttribute("src");
      });
  } else {
    if (img.dataset.blobUrl) {
      URL.revokeObjectURL(img.dataset.blobUrl);
      delete img.dataset.blobUrl;
    }
    img.hidden = true;
    img.removeAttribute("src");
  }
  if (removeBtn) removeBtn.disabled = !customOnly;
}

async function uploadPdfReportLogo() {
  const file = qs("pdfReportLogoFile")?.files?.[0];
  if (!file) return alert("Choose a logo image first (PNG or JPEG).");
  const out = qs("pdfReportSettingsResult");
  const fd = new FormData();
  fd.append("file", file);
  const headers = new Headers(authHeaders());
  headers.delete("Content-Type");
  setStatus("Uploading PDF logo…");
  try {
    const res = await fetch(`${API}/api/reports/pdf-settings/logo`, { method: "POST", headers, body: fd });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || data.message || "Upload failed");
    if (qs("pdfReportLogoFile")) qs("pdfReportLogoFile").value = "";
    updatePdfReportLogoPreview(Boolean(data?.company_logo_available), Boolean(data?.company_logo_custom));
    if (out) out.textContent = data?.message || "PDF logo uploaded.";
    setStatus("PDF logo uploaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("PDF logo upload failed.");
  }
}

async function removePdfReportLogo() {
  const out = qs("pdfReportSettingsResult");
  if (!confirm("Remove the uploaded PDF company logo?")) return;
  setStatus("Removing PDF logo…");
  try {
    const data = await fetchJson(`${API}/api/reports/pdf-settings/logo`, { method: "DELETE" });
    updatePdfReportLogoPreview(Boolean(data?.company_logo_available), Boolean(data?.company_logo_custom));
    if (out) out.textContent = data?.message || "PDF logo removed.";
    setStatus("PDF logo removed.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("PDF logo remove failed.");
  }
}

async function loadPdfReportSettings() {
  const out = qs("pdfReportSettingsResult");
  try {
    const data = await fetchJson(`${API}/api/reports/pdf-settings`);
    pdfReportSitesCache = Array.isArray(data?.sites) ? data.sites : [];
    const companies = Array.isArray(data?.companies) ? data.companies : [];
    const companySel = qs("pdfReportCompanyCode");
    if (companySel) {
      const currentCompany = String(data?.company_code || "");
      companySel.innerHTML = `<option value="">Custom / manual company name</option>${companies.map((c) =>
        `<option value="${escapeHtml(String(c.company_code || ""))}">${escapeHtml(String(c.company_name || c.company_code || ""))}</option>`
      ).join("")}`;
      companySel.value = currentCompany && Array.from(companySel.options).some((o) => o.value === currentCompany)
        ? currentCompany
        : "";
    }
    if (qs("pdfReportCompanyName")) qs("pdfReportCompanyName").value = String(data?.company_name || "");
    const siteSel = qs("pdfReportSiteCode");
    if (siteSel) {
      const currentSite = String(data?.site_code || "");
      const companyFilter = String(data?.company_code || companySel?.value || "").trim();
      const sites = companyFilter
        ? pdfReportSitesCache.filter((s) => String(s.company_code || "") === companyFilter)
        : pdfReportSitesCache;
      siteSel.innerHTML = `<option value="">Custom / manual site name</option>${sites.map((s) =>
        `<option value="${escapeHtml(String(s.site_code || ""))}">${escapeHtml(String(s.site_name || s.site_code || ""))}</option>`
      ).join("")}`;
      siteSel.value = currentSite && Array.from(siteSel.options).some((o) => o.value === currentSite) ? currentSite : "";
    }
    if (qs("pdfReportSiteName")) qs("pdfReportSiteName").value = String(data?.site_name || "");
    updatePdfReportLogoPreview(
      Boolean(data?.company_logo_available),
      Boolean(data?.company_logo_custom)
    );
    const parts = [];
    if (data?.company_name) parts.push(`Company: ${data.company_name}`);
    if (data?.site_name) parts.push(`Site: ${data.site_name}`);
    if (data?.company_logo_custom) parts.push("Custom logo: yes");
    else if (data?.company_logo_available) parts.push("Logo: default");
    if (out) out.textContent = parts.length ? `Current PDF header — ${parts.join(" · ")}` : "No PDF branding configured yet.";
    setStatus("PDF report branding loaded.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("PDF report branding load failed.");
  }
}

async function savePdfReportSettings() {
  const out = qs("pdfReportSettingsResult");
  const company_code = String(qs("pdfReportCompanyCode")?.value || "").trim();
  let company_name = String(qs("pdfReportCompanyName")?.value || "").trim();
  const site_code = String(qs("pdfReportSiteCode")?.value || "").trim();
  let site_name = String(qs("pdfReportSiteName")?.value || "").trim();
  if (company_code && !company_name) {
    const opt = qs("pdfReportCompanyCode")?.selectedOptions?.[0];
    company_name = String(opt?.textContent || company_code).trim();
  }
  if (site_code && !site_name) {
    const opt = qs("pdfReportSiteCode")?.selectedOptions?.[0];
    site_name = String(opt?.textContent || site_code).trim();
  }
  if (!company_name && !company_code && !site_name && !site_code) {
    return alert("Enter a company name and/or site name.");
  }
  setStatus("Saving PDF report branding…");
  try {
    const data = await fetchJson(`${API}/api/reports/pdf-settings`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ company_code, company_name, site_code, site_name }),
    });
    if (qs("pdfReportCompanyName")) qs("pdfReportCompanyName").value = String(data?.company_name || company_name);
    if (qs("pdfReportSiteName")) qs("pdfReportSiteName").value = String(data?.site_name || site_name);
    const parts = [];
    if (data?.company_name || company_name) parts.push(data?.company_name || company_name);
    if (data?.site_name || site_name) parts.push(`Site: ${data?.site_name || site_name}`);
    if (out) out.textContent = parts.length ? `Saved. PDF header will show: ${parts.join(" · ")}` : "Saved.";
    setStatus("PDF report branding saved.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("PDF report branding save failed.");
  }
}

function onPdfReportCompanyCodeChange() {
  const code = String(qs("pdfReportCompanyCode")?.value || "").trim();
  if (code) {
    const opt = qs("pdfReportCompanyCode")?.selectedOptions?.[0];
    const name = String(opt?.textContent || code).trim();
    if (qs("pdfReportCompanyName")) qs("pdfReportCompanyName").value = name;
  }
  const siteSel = qs("pdfReportSiteCode");
  if (!siteSel) return;
  const currentSite = String(siteSel.value || "");
  const sites = code
    ? pdfReportSitesCache.filter((s) => String(s.company_code || "") === code)
    : pdfReportSitesCache;
  siteSel.innerHTML = `<option value="">Custom / manual site name</option>${sites.map((s) =>
    `<option value="${escapeHtml(String(s.site_code || ""))}">${escapeHtml(String(s.site_name || s.site_code || ""))}</option>`
  ).join("")}`;
  if (currentSite && Array.from(siteSel.options).some((o) => o.value === currentSite)) {
    siteSel.value = currentSite;
  } else {
    siteSel.value = "";
    if (qs("pdfReportSiteName")) qs("pdfReportSiteName").value = "";
  }
}

async function onPdfReportSiteCodeChange() {
  const code = String(qs("pdfReportSiteCode")?.value || "").trim();
  if (!code) return;
  const opt = qs("pdfReportSiteCode")?.selectedOptions?.[0];
  const name = String(opt?.textContent || code).trim();
  if (qs("pdfReportSiteName")) qs("pdfReportSiteName").value = name;
}

async function sendSubscriptionNowFromAdmin() {
  const out = qs("smtpSettingsResult");
  const id = Number(qs("smtpSubscriptionId")?.value || 0);
  if (!id) return alert("Select a subscription first.");
  setStatus("Sending subscription now...");
  try {
    const data = await fetchJson(`${API}/api/reports/subscriptions/${id}/send-now`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    if (out) out.textContent = JSON.stringify(data, null, 2);
    setStatus("Subscription sent.");
  } catch (e) {
    if (out) out.textContent = String(e.message || e);
    setStatus("Subscription send failed.");
  }
}

function backupOptionLabel(f) {
  const mb = Number(f?.bytes || 0) / (1024 * 1024);
  const size = Number.isFinite(mb) ? `${mb.toFixed(1)} MB` : "-";
  return `${String(f?.name || "-")} (${size})`;
}

async function loadBackupFiles() {
  const out = qs("backupRestoreResult");
  const sel = qs("backupFileSelect");
  if (!sel) return;
  const current = String(sel.value || "");
  const data = await fetchJson(`${API}/api/admin/backups/list`);
  const files = Array.isArray(data?.files) ? data.files : [];
  const executeEnabled = Boolean(data?.execute_restore_enabled);
  const restartCmdSet = Boolean(data?.restart_command_set);
  sel.innerHTML = files.length
    ? `<option value="">Select backup file...</option>${files
        .map((f) => `<option value="${esc(String(f.name || ""))}">${esc(backupOptionLabel(f))}</option>`)
        .join("")}`
    : `<option value="">No backups found</option>`;
  if (current && Array.from(sel.options).some((o) => o.value === current)) sel.value = current;
  if (out) {
    out.textContent =
      `Backup dir: ${data?.backup_dir || "-"}\n` +
      `DB path: ${data?.db_path || "-"}\n` +
      `Backups found: ${files.length}\n` +
      `Execute restore enabled: ${executeEnabled ? "yes" : "no"}\n` +
      `Restart command set: ${restartCmdSet ? "yes" : "no"}`;
  }
  const execBtn = qs("executeBackupRestoreBtn");
  if (execBtn) execBtn.disabled = !(executeEnabled && restartCmdSet);
}

async function createBackupNow() {
  const out = qs("backupRestoreResult");
  setStatus("Creating manual backup...");
  const data = await fetchJson(`${API}/api/admin/backups/create`, {
    method: "POST",
    body: JSON.stringify({}),
  });
  if (out) out.textContent = JSON.stringify(data, null, 2);
  await loadBackupFiles();
  setStatus("Backup created.");
}

async function previewBackupRestore() {
  const out = qs("backupRestoreResult");
  const backup_name = String(qs("backupFileSelect")?.value || "").trim();
  if (!backup_name) return alert("Select a backup file first.");
  setStatus("Loading restore preview...");
  const data = await fetchJson(`${API}/api/admin/backups/restore/preview`, {
    method: "POST",
    body: JSON.stringify({ backup_name }),
  });
  if (out) out.textContent = JSON.stringify(data, null, 2);
  setStatus("Restore preview ready.");
}

async function stageBackupRestore() {
  const out = qs("backupRestoreResult");
  const backup_name = String(qs("backupFileSelect")?.value || "").trim();
  const notes = String(qs("backupRestoreNotes")?.value || "").trim();
  if (!backup_name) return alert("Select a backup file first.");
  const confirmed = window.confirm(
    `Stage restore plan for ${backup_name}?\n\nThis does not auto-copy yet; it prepares safe restore steps.`
  );
  if (!confirmed) return;
  setStatus("Staging restore plan...");
  const data = await fetchJson(`${API}/api/admin/backups/restore/apply`, {
    method: "POST",
    body: JSON.stringify({ backup_name, confirm_text: "RESTORE", notes }),
  });
  if (out) out.textContent = JSON.stringify(data, null, 2);
  await loadBackupFiles();
  setStatus("Restore plan staged.");
}

async function executeBackupRestoreNow() {
  const out = qs("backupRestoreResult");
  const backup_name = String(qs("backupFileSelect")?.value || "").trim();
  const notes = String(qs("backupRestoreNotes")?.value || "").trim();
  if (!backup_name) return alert("Select a backup file first.");
  const confirmed = window.confirm(
    `Execute restore now for ${backup_name}?\n\nThis will restart the API process immediately.`
  );
  if (!confirmed) return;
  setStatus("Executing restore + restart...");
  const data = await fetchJson(`${API}/api/admin/backups/restore/execute`, {
    method: "POST",
    body: JSON.stringify({ backup_name, confirm_text: "RESTORE_NOW", notes }),
  });
  if (out) out.textContent = JSON.stringify(data, null, 2);
  setStatus("Restore execute requested. Reconnect after restart.");
}

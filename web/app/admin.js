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
    // The company's real sites (api/utils/sites.js), not whatever was typed on other users.
    try {
      const siteData = await fetchJson(`${API}/api/auth/sites`);
      knownLocations.clear();
      (siteData.sites || []).forEach((s) => knownLocations.add(s));
      if (qs("adminSitesHint")) qs("adminSitesHint").textContent = `Sites this person may use, comma separated. Leave empty for all sites. Sites: ${(siteData.sites || []).join(", ")}. Screens are chosen under Sections, not here.`;
    } catch {}
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
      tr.innerHTML = `<td>${escapeHtml(r.username)}</td><td>${escapeHtml(r.full_name || "")}</td><td>${escapeHtml(r.department || "")}</td><td>${escapeHtml(rolesText)}</td><td>${escapeHtml(locText)}</td><td>${r.active ? "yes" : "no"}</td><td>${r.has_password ? "yes" : "no"}</td><td>${r.has_pin ? "yes" : "no"}</td><td>${r.microsoft_email ? `${escapeHtml(r.microsoft_email)} <span class="pill ${r.microsoft_linked ? "pill-green" : "pill-gray"}">${r.microsoft_linked ? "Linked" : "Invited"}</span>${r.microsoft_only ? ' <span class="pill pill-blue">Microsoft only</span>' : ""}${r.roles_from_groups ? ' <span class="pill pill-blue">Roles from groups</span>' : ""}` : ""}</td>`;
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
        if (qs("adminMsEmail")) qs("adminMsEmail").value = r.microsoft_email || "";
        if (qs("adminMsEmailNote")) {
          qs("adminMsEmailNote").textContent = r.microsoft_linked
            ? "Linked: this person has signed in with Microsoft. Changing the email unlinks their account."
            : r.microsoft_email ? "Invited: they can sign in with Microsoft using this work email." : "Invites this person to sign in with Microsoft. Leave empty for no Microsoft sign-in.";
        }
        if (qs("adminMsOnly")) qs("adminMsOnly").checked = Boolean(r.microsoft_only);
        if (qs("adminMsGroups")) qs("adminMsGroups").checked = Boolean(r.roles_from_groups);
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
  if (qs("adminMsEmail") && !qs("adminMsEmail").closest(".hidden")) {
    body.microsoft_email = String(qs("adminMsEmail").value || "").trim();
    body.microsoft_only = qs("adminMsOnly")?.checked === true;
    body.roles_from_groups = qs("adminMsGroups")?.checked === true;
    if (body.microsoft_only && password) return alert("A Microsoft-only user has no password. Clear the password, or untick 'Microsoft sign-in only'.");
    if (body.microsoft_only) body.issue_setup_code = false;
  }
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
    if (hideWhenForbidden(e, "adminTelematicsCard")) return;
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

/** Admin → Microsoft sign-in (api/routes/microsoftSignin.routes.js). */
async function loadMsSignin() {
  const card = qs("msSigninCard");
  if (!card || card.classList.contains("hidden")) return;
  const data = await fetchJson(`${API}/api/sso/microsoft/settings`);
  const s = data.settings || {};
  qs("msTenantId").value = s.tenant_id || "";
  qs("msClientId").value = s.client_id || "";
  qs("msClientSecret").value = "";
  qs("msClientSecret").placeholder = s.has_secret ? "Saved. Leave blank to keep it" : "Paste the secret's Value";
  qs("msSecretExpires").value = s.secret_expires_on || "";
  qs("msEnabled").checked = Boolean(s.enabled);
  if (qs("msOpenToCompany")) qs("msOpenToCompany").checked = Boolean(s.open_to_company);
  msRoleChoices = Array.isArray(data.roles) ? data.roles : msRoleChoices;
  renderMsGroups(Array.isArray(data.groups) ? data.groups : []);
  qs("msRedirectUri").textContent = data.redirect_uri || "";
  const days = data.secret_days_left;
  const state = !s.configured
    ? ["is-off", "Not set up yet. Follow the steps below with your IT administrator."]
    : !s.enabled
      ? ["is-off", "Set up but switched off: only the Test sign-in works."]
      : ["is-ok", `Switched on${s.open_to_company ? " and open to the whole company" : ""}. ${data.invited} invited, ${data.linked} already signed in with Microsoft${data.ms_only ? `, ${data.ms_only} Microsoft-only` : ""}${data.roles_from_groups ? `, ${data.roles_from_groups} with roles from groups` : ""}.`];
  const warn = days != null && days <= 30
    ? `<div class="ms-signin-warn">${days < 0 ? "The client secret has run out: Microsoft sign-in stops working. Create a new secret in Entra and save it here." : `The client secret runs out in ${days} days. Create a new one in Entra and save it here before then.`}</div>`
    : "";
  qs("msSigninState").innerHTML = `<div class="ms-signin-pill ${state[0]}">${escapeHtml(state[1])}</div>${warn}`;
}

let msRoleChoices = [];

/** One editable line per linked Microsoft group. */
function msGroupRow(g = {}) {
  const row = document.createElement("div");
  row.className = "ms-group-row";
  const roles = new Set(g.roles || []);
  row.innerHTML = `
    <label>Group Object ID<input class="ms-g-id w-full" value="${escapeHtml(g.group_id || "")}" placeholder="1a2b3c4d-1234-5678-9abc-1234567890ab" autocomplete="off" spellcheck="false" /></label>
    <label>Name (for you)<input class="ms-g-label w-full" value="${escapeHtml(g.label || "")}" placeholder="e.g. Workshop technicians" maxlength="80" /></label>
    <label>IronLog roles<select class="ms-g-roles w-full" multiple size="4">${msRoleChoices.map((r) => `<option value="${escapeHtml(r)}"${roles.has(r) ? " selected" : ""}>${escapeHtml(r)}</option>`).join("")}</select></label>
    <label>Sites (empty = all)<input class="ms-g-sites w-full" value="${escapeHtml((g.sites || []).join(", "))}" placeholder="e.g. main" /></label>
    <button type="button" class="btn btn-secondary btn-sm ms-g-remove" aria-label="Remove group">Remove</button>`;
  row.querySelector(".ms-g-remove").addEventListener("click", () => row.remove());
  return row;
}

function renderMsGroups(groups) {
  const list = qs("msGroupsList");
  if (!list) return;
  list.innerHTML = "";
  groups.forEach((g) => list.appendChild(msGroupRow(g)));
}

function collectMsGroups() {
  return Array.from(document.querySelectorAll("#msGroupsList .ms-group-row")).map((row) => ({
    group_id: row.querySelector(".ms-g-id").value.trim(),
    label: row.querySelector(".ms-g-label").value.trim(),
    roles: Array.from(row.querySelector(".ms-g-roles").selectedOptions).map((o) => o.value),
    sites: row.querySelector(".ms-g-sites").value.split(",").map((x) => x.trim()).filter(Boolean),
  })).filter((g) => g.group_id || g.label || g.roles.length);
}

async function saveMsSignin() {
  const out = qs("msSigninResult");
  try {
    await fetchJson(`${API}/api/sso/microsoft/settings`, {
      method: "PUT",
      body: JSON.stringify({
        tenant_id: qs("msTenantId").value,
        client_id: qs("msClientId").value,
        client_secret: qs("msClientSecret").value,
        secret_expires_on: qs("msSecretExpires").value,
        enabled: qs("msEnabled").checked,
        open_to_company: qs("msOpenToCompany")?.checked === true,
        groups: collectMsGroups(),
      }),
    });
    await loadMsSignin();
    if (out) out.textContent = "Saved.";
  } catch (e) {
    if (out) out.textContent = `Not saved: ${e.message || e}`;
  }
}

/** Admin → API keys for other systems (api/routes/apiKeys.routes.js). */
async function loadApiKeys() {
  const card = qs("apiKeysCard");
  if (!card || card.classList.contains("hidden")) return;
  const [data, cat] = await Promise.all([
    fetchJson(`${API}/api/api-keys`),
    fetchJson(`${API}/api/api-keys/catalogue`),
  ]);
  const labels = Object.fromEntries((cat.scopes || []).map((s) => [s.scope, s.label]));
  const box = qs("apiKeyScopes");
  if (box && !box.children.length) {
    box.innerHTML = (cat.scopes || []).map((s) => `<label><input type="checkbox" value="${escapeHtml(s.scope)}" /> <span><b>${escapeHtml(s.label)}</b><br/><small class="muted">${s.endpoints.length} address${s.endpoints.length === 1 ? "" : "es"}</small></span></label>`).join("");
  }
  const keys = data.keys || [];
  const statusPill = (k) => `<span class="pill ${k.status === "active" ? "pill-green" : "pill-gray"}">${k.status === "active" ? "Active" : k.status === "expired" ? "Expired" : "Cancelled"}</span>`;
  qs("apiKeysList").innerHTML = keys.length
    ? `<div class="table-wrap"><table class="gridTable"><thead><tr><th>Name</th><th>Key</th><th>Can read</th><th>Sites</th><th>Last used</th><th>Status</th><th></th></tr></thead><tbody>${keys.map((k) => `
      <tr><td>${escapeHtml(k.name)}<br/><small class="muted">by ${escapeHtml(k.created_by || "—")}, ${escapeHtml(String(k.created_at || "").slice(0, 10))}${k.expires_on ? ` · until ${escapeHtml(k.expires_on)}` : ""}</small></td>
      <td><code>${escapeHtml(k.key_prefix)}…</code></td>
      <td>${k.scopes.map((s) => escapeHtml(labels[s] || s)).join(", ")}</td>
      <td>${k.sites ? escapeHtml(k.sites.join(", ")) : "All"}</td>
      <td>${k.last_used_at ? `${escapeHtml(String(k.last_used_at).slice(0, 16))}<br/><small class="muted">${Number(k.use_count || 0).toLocaleString()} calls</small>` : "Never"}</td>
      <td>${statusPill(k)}</td>
      <td>${k.status === "active" ? `<button type="button" class="btn btn-secondary btn-sm api-key-cancel" data-id="${k.id}" data-name="${escapeHtml(k.name)}">Cancel</button>` : ""}</td></tr>`).join("")}</tbody></table></div>`
    : `<p class="muted small">No API keys yet.</p>`;
}

async function createApiKeyFromForm() {
  const scopes = Array.from(document.querySelectorAll("#apiKeyScopes input:checked")).map((c) => c.value);
  const out = qs("apiKeyNew");
  try {
    const res = await fetchJson(`${API}/api/api-keys`, {
      method: "POST",
      body: JSON.stringify({ name: qs("apiKeyName").value, scopes, sites: qs("apiKeySites").value, expires_on: qs("apiKeyExpires").value }),
    });
    out.innerHTML = `<b>Copy this key now: it is shown only once.</b> Give it to the people setting up ${escapeHtml(res.row.name)}; they send it as <code>Authorization: Bearer &lt;key&gt;</code>.`
      + `<code id="apiKeyNewValue">${escapeHtml(res.key)}</code><button type="button" class="btn btn-secondary btn-sm" id="apiKeyCopyBtn">Copy key</button>`;
    out.classList.remove("hidden");
    qs("apiKeyCopyBtn")?.addEventListener("click", async () => {
      try { await navigator.clipboard.writeText(res.key); qs("apiKeyCopyBtn").textContent = "Copied"; } catch { /* select by hand */ }
    });
    qs("apiKeyName").value = "";
    document.querySelectorAll("#apiKeyScopes input:checked").forEach((c) => { c.checked = false; });
    if (qs("apiKeysMsg")) qs("apiKeysMsg").textContent = "";
    await loadApiKeys();
  } catch (e) {
    out.classList.add("hidden");
    if (qs("apiKeysMsg")) qs("apiKeysMsg").textContent = `Not created: ${e.message || e}`;
  }
}

function wireApiKeys() {
  qs("apiKeysRefreshBtn")?.addEventListener("click", () => loadApiKeys().catch((e) => setStatus("API keys error: " + e.message)));
  qs("apiKeyCreateBtn")?.addEventListener("click", () => createApiKeyFromForm());
  qs("apiKeysList")?.addEventListener("click", async (ev) => {
    const b = ev.target?.closest?.(".api-key-cancel");
    if (!b) return;
    if (!confirm(`Cancel the key "${b.dataset.name}"? The system using it stops getting data straight away. This cannot be undone.`)) return;
    try {
      await fetchJson(`${API}/api/api-keys/${b.dataset.id}/cancel`, { method: "POST", body: "{}" });
      qs("apiKeyNew")?.classList.add("hidden");
      await loadApiKeys();
    } catch (e) {
      if (qs("apiKeysMsg")) qs("apiKeysMsg").textContent = e.message || String(e);
    }
  });
}

/** Admin → Webhooks: messages to other systems when something happens. */
async function loadWebhooks() {
  const card = qs("webhooksCard");
  if (!card || card.classList.contains("hidden")) return;
  const data = await fetchJson(`${API}/api/webhooks`);
  const labels = Object.fromEntries((data.events || []).map((e) => [e.event, e.label]));
  const box = qs("webhookEvents");
  if (box && !box.children.length) {
    box.innerHTML = (data.events || []).map((e) => `<label><input type="checkbox" value="${escapeHtml(e.event)}" /> <span><b>${escapeHtml(e.label)}</b><br/><small class="muted">${escapeHtml(e.description)}</small></span></label>`).join("");
  }
  const hooks = data.webhooks || [];
  const state = (h) => !h.active ? `<span class="pill pill-gray">Paused</span>`
    : h.failures_in_row > 0 ? `<span class="pill pill-red">Failing (${h.failures_in_row})</span>`
      : `<span class="pill pill-green">On</span>`;
  qs("webhooksList").innerHTML = hooks.length
    ? `<div class="table-wrap"><table class="gridTable"><thead><tr><th>Name</th><th>Address</th><th>Sends</th><th>Sites</th><th>Last sent</th><th>Status</th><th></th></tr></thead><tbody>${hooks.map((h) => `
      <tr><td>${escapeHtml(h.name)}<br/><small class="muted">by ${escapeHtml(h.created_by || "—")}, ${escapeHtml(String(h.created_at || "").slice(0, 10))} · secret <code>${escapeHtml(h.secret_prefix)}…</code></small></td>
      <td class="webhook-url">${escapeHtml(h.url)}</td>
      <td>${h.events.map((e) => escapeHtml(labels[e] || e)).join(", ")}</td>
      <td>${h.sites ? escapeHtml(h.sites.join(", ")) : "All"}</td>
      <td>${h.last_success_at ? escapeHtml(String(h.last_success_at).slice(0, 16)) : "Never"}${h.last_failure_at ? `<br/><small class="muted">last failure ${escapeHtml(String(h.last_failure_at).slice(0, 16))}</small>` : ""}</td>
      <td>${state(h)}</td>
      <td class="webhook-actions">
        <button type="button" class="btn btn-secondary btn-sm" data-wh-act="test" data-id="${h.id}">Send test</button>
        <button type="button" class="btn btn-secondary btn-sm" data-wh-act="log" data-id="${h.id}" data-name="${escapeHtml(h.name)}">Messages</button>
        <button type="button" class="btn btn-secondary btn-sm" data-wh-act="${h.active ? "pause" : "resume"}" data-id="${h.id}">${h.active ? "Pause" : "Resume"}</button>
        <button type="button" class="btn btn-secondary btn-sm" data-wh-act="secret" data-id="${h.id}" data-name="${escapeHtml(h.name)}">New secret</button>
        <button type="button" class="btn btn-secondary btn-sm" data-wh-act="delete" data-id="${h.id}" data-name="${escapeHtml(h.name)}">Delete</button>
      </td></tr>`).join("")}</tbody></table></div>`
    : `<p class="muted small">No webhooks yet.</p>`;
}

function showWebhookSecret(name, secret, intro) {
  const out = qs("webhookNew");
  out.innerHTML = `<b>Copy this secret now: it is shown only once.</b> ${escapeHtml(intro)} The people setting up ${escapeHtml(name)} use it to check each message is really from IronLog.`
    + `<code id="webhookSecretValue">${escapeHtml(secret)}</code><button type="button" class="btn btn-secondary btn-sm" id="webhookCopyBtn">Copy secret</button>`;
  out.classList.remove("hidden");
  qs("webhookCopyBtn")?.addEventListener("click", async () => {
    try { await navigator.clipboard.writeText(secret); qs("webhookCopyBtn").textContent = "Copied"; } catch { /* select by hand */ }
  });
}

async function createWebhookFromForm() {
  const events = Array.from(document.querySelectorAll("#webhookEvents input:checked")).map((c) => c.value);
  try {
    const res = await fetchJson(`${API}/api/webhooks`, {
      method: "POST",
      body: JSON.stringify({ name: qs("webhookName").value, url: qs("webhookUrl").value, sites: qs("webhookSites").value, events }),
    });
    showWebhookSecret(res.webhook.name, res.secret, "Webhook added.");
    qs("webhookName").value = "";
    qs("webhookUrl").value = "";
    document.querySelectorAll("#webhookEvents input:checked").forEach((c) => { c.checked = false; });
    if (qs("webhooksMsg")) qs("webhooksMsg").textContent = "Press Send test to check the other system receives it.";
    await loadWebhooks();
  } catch (e) {
    qs("webhookNew")?.classList.add("hidden");
    if (qs("webhooksMsg")) qs("webhooksMsg").textContent = `Not added: ${e.message || e}`;
  }
}

async function showWebhookDeliveries(id, name) {
  const box = qs("webhookDeliveries");
  const data = await fetchJson(`${API}/api/webhooks/${id}/deliveries`);
  const rows = data.deliveries || [];
  const stateOf = (d) => d.status === "delivered" ? `<span class="pill pill-green">Delivered</span>`
    : d.status === "failed" ? `<span class="pill pill-red">Failed</span>` : `<span class="pill pill-gray">Waiting (try ${d.attempts + 1})</span>`;
  box.innerHTML = `<h4>Messages to ${escapeHtml(name)}</h4>` + (rows.length
    ? `<div class="table-wrap"><table class="gridTable"><thead><tr><th>When</th><th>Event</th><th>Record</th><th>Status</th><th>Answer</th><th></th></tr></thead><tbody>${rows.map((d) => `
      <tr><td>${escapeHtml(String(d.created_at || "").slice(0, 16))}</td><td><code>${escapeHtml(d.event)}</code></td><td>${escapeHtml(d.record_key || "")}</td><td>${stateOf(d)}</td>
      <td>${d.last_error ? escapeHtml(d.last_error) : d.last_status_code ? escapeHtml(String(d.last_status_code)) : ""}${d.status === "pending" && d.attempts ? `<br/><small class="muted">next try ${escapeHtml(String(d.next_attempt_at).slice(0, 16))}</small>` : ""}</td>
      <td>${d.status !== "delivered" ? `<button type="button" class="btn btn-secondary btn-sm" data-wh-act="retry" data-id="${id}" data-did="${d.id}" data-name="${escapeHtml(name)}">Send again</button>` : ""}</td></tr>`).join("")}</tbody></table></div>`
    : `<p class="muted small">No messages yet.</p>`);
  box.classList.remove("hidden");
}

async function webhookAction(b) {
  const id = b.dataset.id;
  const name = b.dataset.name || "";
  const msg = (t) => { if (qs("webhooksMsg")) qs("webhooksMsg").textContent = t; };
  const act = b.dataset.whAct;
  if (act === "test") {
    b.disabled = true;
    try {
      const r = await fetchJson(`${API}/api/webhooks/${id}/test`, { method: "POST", body: "{}" });
      msg(r.ok ? `Test delivered (the system answered ${r.status_code}).` : `Test not delivered: ${r.error || r.status}`);
    } finally { b.disabled = false; }
  } else if (act === "log") {
    await showWebhookDeliveries(id, name);
    return;
  } else if (act === "retry") {
    await fetchJson(`${API}/api/webhooks/${id}/deliveries/${b.dataset.did}/retry`, { method: "POST", body: "{}" });
    await showWebhookDeliveries(id, name);
  } else if (act === "pause" || act === "resume") {
    await fetchJson(`${API}/api/webhooks/${id}`, { method: "PATCH", body: JSON.stringify({ active: act === "resume" }) });
    msg(act === "pause" ? "Paused: nothing is sent until you resume it." : "Resumed.");
  } else if (act === "secret") {
    if (!confirm(`Make a new secret for "${name}"? The old one stops working straight away, so the other system must be given the new one.`)) return;
    const r = await fetchJson(`${API}/api/webhooks/${id}/secret`, { method: "POST", body: "{}" });
    showWebhookSecret(name, r.secret, "New secret made.");
  } else if (act === "delete") {
    if (!confirm(`Delete the webhook "${name}"? It stops at once and its message history is removed.`)) return;
    await fetchJson(`${API}/api/webhooks/${id}`, { method: "DELETE" });
    qs("webhookDeliveries")?.classList.add("hidden");
    msg("Deleted.");
  }
  await loadWebhooks();
}

function wireWebhooks() {
  qs("webhooksRefreshBtn")?.addEventListener("click", () => loadWebhooks().catch((e) => setStatus("Webhooks error: " + e.message)));
  qs("webhookCreateBtn")?.addEventListener("click", () => createWebhookFromForm());
  for (const id of ["webhooksList", "webhookDeliveries"]) {
    qs(id)?.addEventListener("click", (ev) => {
      const b = ev.target?.closest?.("[data-wh-act]");
      if (!b) return;
      webhookAction(b).catch((e) => { if (qs("webhooksMsg")) qs("webhooksMsg").textContent = e.message || String(e); });
    });
  }
}

function wireMsSignin() {
  qs("msSigninRefreshBtn")?.addEventListener("click", () => loadMsSignin().catch((e) => setStatus("Microsoft sign-in error: " + e.message)));
  qs("msSaveBtn")?.addEventListener("click", () => saveMsSignin());
  qs("msAddGroupBtn")?.addEventListener("click", () => qs("msGroupsList")?.appendChild(msGroupRow()));
  // Opens Microsoft in a new tab; the result page there says whether it worked. Nobody is signed in.
  qs("msTestBtn")?.addEventListener("click", () => window.open(`${API}/api/sso/microsoft/start?test=1`, "_blank", "noopener"));
  qs("msCopyRedirectBtn")?.addEventListener("click", async () => {
    const text = qs("msRedirectUri")?.textContent || "";
    try {
      await navigator.clipboard.writeText(text);
      qs("msCopyRedirectBtn").textContent = "Copied";
    } catch {
      window.prompt("Copy this address:", text);
    }
  });
}

/** Admin → Company profile: details used on report headers and June's site. */
async function loadCompanyProfile() {
  const card = qs("companyProfileCard");
  if (!card) return;
  const data = await fetchJson(`${API}/api/company/profile`);
  const p = data?.profile || {};
  card.querySelectorAll("[data-cp]").forEach((el) => { el.value = String(p[el.dataset.cp] ?? ""); });
  const list = qs("cpTimezoneList");
  if (list && !list.children.length && typeof Intl.supportedValuesOf === "function") {
    list.innerHTML = Intl.supportedValuesOf("timeZone").map((z) => `<option value="${escapeHtml(z)}"></option>`).join("");
  }
  const out = qs("companyProfileResult");
  if (out) out.textContent = p.header_prefix ? `Report headers show "${p.header_prefix} · IronLog".` : "Report headers show \"IronLog\" until a short name or company name is set.";
}

async function saveCompanyProfile() {
  const card = qs("companyProfileCard");
  const out = qs("companyProfileResult");
  const body = {};
  card.querySelectorAll("[data-cp]").forEach((el) => { body[el.dataset.cp] = el.value; });
  try {
    await fetchJson(`${API}/api/company/profile`, { method: "PUT", body: JSON.stringify(body) });
    await loadCompanyProfile();
    if (out) out.textContent = `Saved. ${out.textContent}`;
  } catch (e) {
    if (out) out.textContent = `Not saved: ${e.message || e}`;
  }
}

/** Admin → Off-site backups and system health (api/utils/autoBackup.js). */
const BK_KIND = { nightly: "Nightly", manual: "Manual", restore_test: "Restore test" };
const bkSize = (b) => (b == null ? "—" : b > 1e9 ? `${(b / 1e9).toFixed(2)} GB` : `${(b / 1e6).toFixed(1)} MB`);
const bkWhen = (t) => (t ? String(t).slice(0, 16) : "—");

/** One bar per disk (server disk, and the backup volume when copies live there) plus what IronLog keeps on them. */
function diskGaugeHtml(disks = [], usage = []) {
  if (!disks.length) return `<p class="muted">Disk space could not be read on this server.</p>`;
  const bars = disks.map((disk) => {
    const used = Math.max(0, Math.min(100, Number(disk.used_pct) || 0));
    const state = disk.level === "critical" ? "is-critical" : disk.level === "warn" ? "is-warn" : "is-ok";
    const name = escapeHtml(disk.name || "Server disk");
    return `<div class="disk-gauge-one ${state}">
      <div class="disk-gauge-head"><strong>${name}</strong><span>${used}% used · ${escapeHtml(String(disk.free_gb))} GB free of ${escapeHtml(String(disk.total_gb))} GB</span></div>
      <div class="disk-gauge-bar" role="meter" aria-label="${name} used" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${used}"><span style="width:${used}%"></span></div>
    </div>`;
  }).join("");
  const rows = (usage || []).map((u) => `<li><span>${escapeHtml(u.label)}${u.files > 1 ? ` <span class="muted">(${u.files}${u.partial ? "+" : ""} files)</span>` : ""}</span><strong>${escapeHtml(bkSize(u.bytes))}</strong></li>`).join("");
  return `<div class="disk-gauge">
      ${bars}
      <p class="muted disk-gauge-note">Admins get a warning at the top of the screen when less than 10% (or 5 GB) is free.</p>
      ${rows ? `<div class="disk-usage-h">What IronLog keeps on this server</div><ul class="disk-usage">${rows}</ul>` : ""}
    </div>`;
}

const DISK_RANK = { ok: 0, warn: 1, critical: 2 };
const worstOfDisks = (disks = []) => disks.reduce((w, d) => (!w || DISK_RANK[d.level] > DISK_RANK[w.level] ? d : w), null);

async function loadBackupHealth() {
  const card = qs("offsiteBackupCard");
  if (!card) return;
  const data = await fetchJson(`${API}/api/admin/backups/status`);
  const s = data.settings || {};
  const hour = qs("bkHour");
  if (hour && !hour.options.length) hour.innerHTML = Array.from({ length: 24 }, (_, h) => `<option value="${h}">${String(h).padStart(2, "0")}:00</option>`).join("");
  qs("bkUrl").value = s.webdav_url || "";
  qs("bkUser").value = s.webdav_user || "";
  qs("bkPassword").value = "";
  qs("bkPassword").placeholder = s.has_password ? "Saved. Leave blank to keep it" : "Storage password";
  qs("bkFolder").value = s.remote_folder || "";
  hour.value = String(s.run_hour ?? 2);
  qs("bkKeepDays").value = String(s.keep_remote_days ?? 30);
  qs("bkEnabled").checked = s.enabled !== false;
  qs("bkFiles").checked = s.include_files !== false;
  const lb = data.last_good_backup;
  const lo = data.last_offsite;
  const lr = data.last_restore_test;
  const srv = data.server || {};
  const lines = [
    ["Last good backup", lb ? `${bkWhen(lb.finished_at)} UTC, ${bkSize(lb.db_bytes)}` : "None yet"],
    ["Last off-site copy", lo ? `${bkWhen(lo.finished_at)} UTC` : s.offsite_configured ? "None yet" : "Storage not set up"],
    ["Last restore test", lr ? `${bkWhen(lr.finished_at)} UTC, ${lr.status === "ok" ? "passed" : "failed"}` : "None yet"],
    ["Database size", bkSize(srv.db_bytes)],
    ["Server running for", `${srv.uptime_hours ?? "—"} h`],
  ];
  if (srv.files_waiting) lines.push(["Files waiting to go off-site", String(srv.files_waiting)]);
  const warn = (data.warnings || []).map((w) => `<li>${escapeHtml(w)}</li>`).join("");
  qs("backupHealthSummary").innerHTML = `
    <div class="backup-health-state ${data.ok ? "is-ok" : "is-warn"}">${data.ok ? "All good: backups are running and copied off-site." : "Needs attention"}</div>
    ${warn ? `<ul class="backup-health-warnings">${warn}</ul>` : ""}
    <dl class="backup-health-grid">${lines.map(([k, v]) => `<div><dt>${escapeHtml(k)}</dt><dd>${escapeHtml(v)}</dd></div>`).join("")}</dl>
    ${diskGaugeHtml(srv.disks || (srv.disk ? [srv.disk] : []), srv.usage)}`;
  if (typeof renderDiskBanner === "function") renderDiskBanner(worstOfDisks(srv.disks || (srv.disk ? [srv.disk] : [])));
  qs("bkRunsBody").innerHTML = (data.recent || []).map((r) => `<tr>
      <td>${escapeHtml(bkWhen(r.started_at))}</td><td>${escapeHtml(BK_KIND[r.kind] || r.kind)}</td>
      <td><span class="pill ${r.status === "ok" ? "pill-green" : r.status === "failed" ? "pill-red" : "pill-gray"}">${escapeHtml(r.status === "ok" ? "OK" : r.status === "failed" ? "Failed" : "Running")}</span></td>
      <td>${escapeHtml(bkSize(r.db_bytes))}</td><td>${escapeHtml(r.offsite || "—")}</td>
      <td>${r.files_uploaded || 0}${r.files_pending ? ` (+${r.files_pending} next run)` : ""}</td><td>${escapeHtml(r.error || "")}</td></tr>`).join("")
    || `<tr><td colspan="7" class="muted">No runs yet.</td></tr>`;
}

/** Waits (up to 10 minutes) for the job just started to finish; returns its run, or null. */
async function waitForBackupRun(kinds) {
  for (let i = 0; i < 120; i++) {
    await new Promise((r) => setTimeout(r, 5000));
    const st = await fetchJson(`${API}/api/admin/backups/status`).catch(() => null);
    if (st && !st.running) return (st.recent || []).find((r) => kinds.includes(r.kind)) || null;
  }
  return null;
}

function backupSettingsBody() {
  return {
    webdav_url: qs("bkUrl").value,
    webdav_user: qs("bkUser").value,
    webdav_password: qs("bkPassword").value,
    remote_folder: qs("bkFolder").value,
    run_hour: Number(qs("bkHour").value),
    keep_remote_days: Number(qs("bkKeepDays").value),
    enabled: qs("bkEnabled").checked,
    include_files: qs("bkFiles").checked,
  };
}

async function backupAction(btnId, label, fn) {
  const btn = qs(btnId);
  const out = qs("bkResult");
  btn.disabled = true;
  if (out) out.textContent = label;
  try {
    const msg = await fn();
    if (out) out.textContent = msg;
  } catch (e) {
    if (out) out.textContent = e.message || String(e);
  } finally {
    btn.disabled = false;
    loadBackupHealth().catch(() => {});
  }
}

function wireBackupHealth() {
  qs("refreshBackupHealthBtn")?.addEventListener("click", () => loadBackupHealth().catch((e) => setStatus("Backup status error: " + e.message)));
  qs("bkSaveBtn")?.addEventListener("click", () => backupAction("bkSaveBtn", "Saving…", async () => {
    await fetchJson(`${API}/api/admin/backups/settings`, { method: "PUT", body: JSON.stringify(backupSettingsBody()) });
    return "Saved.";
  }));
  qs("bkTestBtn")?.addEventListener("click", () => backupAction("bkTestBtn", "Connecting…", async () => {
    const b = backupSettingsBody();
    const r = await fetchJson(`${API}/api/admin/backups/test-connection`, { method: "POST", body: JSON.stringify(b.webdav_password ? b : {}) });
    return r.message || "Connected.";
  }));
  qs("bkRunBtn")?.addEventListener("click", () => backupAction("bkRunBtn", "Backing up… the first time can take several minutes.", async () => {
    await fetchJson(`${API}/api/admin/backups/run-now`, { method: "POST", body: "{}" });
    const run = await waitForBackupRun(["manual"]);
    if (!run) return "Still running. The result will appear under Recent runs.";
    if (run.status !== "ok") return `Backup failed: ${run.error || "unknown problem"}`;
    return run.offsite === "uploaded"
      ? `Done: copy checked and sent off-site${run.files_uploaded ? `, plus ${run.files_uploaded} files` : ""}${run.files_pending ? ` (${run.files_pending} more go tonight)` : ""}.`
      : "Done: copy checked and kept on this server (off-site storage not set up).";
  }));
  qs("bkRestoreTestBtn")?.addEventListener("click", () => backupAction("bkRestoreTestBtn", "Downloading the newest off-site copy…", async () => {
    await fetchJson(`${API}/api/admin/backups/restore-test`, { method: "POST", body: "{}" });
    const run = await waitForBackupRun(["restore_test"]);
    if (!run) return "Still running. The result will appear under Recent runs.";
    return run.status === "ok" ? "Restore test passed: the newest off-site copy opens and is complete." : `Restore test failed: ${run.error || "unknown problem"}`;
  }));
}

/** Admin → Licence: modules per site, end date, usage against the bands, and installing a key. */
const LICENCE_MODE_TEXT = {
  legacy: "Existing installation: everything is switched on. No licence key installed yet.",
  trial: "Trial: everything is switched on until the trial ends.",
  licensed: "Licensed.",
  grace: "Licence ended: in the grace period.",
  read_only: "Licence ended: IronLog is read-only.",
};

async function loadLicenceCard() {
  if (!qs("licenceCard")) return;
  const s = await fetchJson(`${API}/api/licence/state`);
  const facts = [
    ["Status", LICENCE_MODE_TEXT[s.mode] || s.mode],
    s.licensee ? ["Licensed to", s.licensee] : null,
    s.package ? ["Package", s.package] : null,
    s.expires_at ? [s.mode === "trial" ? "Trial ends" : "Licence ends", `${s.expires_at}${s.days_left != null && s.days_left >= 0 ? ` (${s.days_left} days)` : ""}`] : null,
    s.read_only_from && s.mode !== "legacy" ? ["Read-only from", s.read_only_from] : null,
  ].filter(Boolean);
  const warn = (s.warnings || []).map((w) => `<li>${escapeHtml(w)}</li>`).join("");
  qs("licenceSummary").innerHTML = `
    <div class="backup-health-state ${["grace", "read_only"].includes(s.mode) ? "is-warn" : "is-ok"}">${escapeHtml(LICENCE_MODE_TEXT[s.mode] || s.mode)}</div>
    ${warn ? `<ul class="backup-health-warnings">${warn}</ul>` : ""}
    <dl class="backup-health-grid">${facts.map(([k, v]) => `<div><dt>${escapeHtml(k)}</dt><dd>${escapeHtml(v)}</dd></div>`).join("")}</dl>`;
  const on = new Set(s.modules || []);
  qs("licenceModules").innerHTML = `<span class="pill pill-green">Core: machines, work orders, maintenance, inspections, technician portal</span>` +
    Object.entries(s.module_labels || {}).map(([k, label]) => `<span class="pill ${on.has(k) ? "pill-green" : "pill-gray"}">${on.has(k) ? "" : "Not included: "}${escapeHtml(label)}</span>`).join("");
  const u = await fetchJson(`${API}/api/licence/usage`);
  qs("licenceUsageBody").innerHTML = (u.sites || []).map((r) => `<tr>
      <td>${escapeHtml(r.site_code)}</td><td>${escapeHtml(r.package || "—")}</td>
      <td>${r.active_assets}${r.over_band ? ' <span class="pill pill-orange">over band</span>' : ""}</td><td>${r.asset_limit ?? "No limit"}</td></tr>`).join("")
    || `<tr><td colspan="4" class="muted">No machines yet.</td></tr>`;
  qs("licenceUsers").textContent = `${u.active_users} active user${u.active_users === 1 ? "" : "s"}.`;
}

function wireLicenceCard() {
  qs("refreshLicenceBtn")?.addEventListener("click", () => loadLicenceCard().catch((e) => setStatus("Licence error: " + e.message)));
  qs("licenceUsageCsvBtn")?.addEventListener("click", () =>
    downloadAuthedFile(`${API}/api/licence/usage.csv`, "ironlog-usage.csv").catch((e) => setStatus("Usage report error: " + e.message))
  );
  qs("licenceInstallBtn")?.addEventListener("click", async () => {
    const out = qs("licenceResult");
    const key = String(qs("licenceKeyInput")?.value || "").trim();
    if (!key) { if (out) out.textContent = "Paste the licence key first."; return; }
    try {
      const r = await fetchJson(`${API}/api/licence/install`, { method: "POST", body: JSON.stringify({ key }) });
      qs("licenceKeyInput").value = "";
      if (out) out.textContent = `Licence installed for ${r.state?.licensee || "this company"}, ending ${r.state?.expires_at || "—"}.`;
      LICENCE_STATE = null;
      licenceLoadedFor = "";
      applyRoleVisibility();
      await loadLicenceCard();
    } catch (e) {
      if (out) out.textContent = `Not installed: ${e.message || e}`;
    }
  });
}

/** Start-up: User admin, SMTP, push, PDF settings and backup controls. Called once from init() in init.js. */
function wireAdminControls() {
  wireBackupHealth();
  wireMsSignin();
  wireApiKeys();
  wireWebhooks();
  wireLicenceCard();
  wireLicenceIssuer();
  qs("loadCompanyProfileBtn")?.addEventListener("click", () =>
    loadCompanyProfile().catch((e) => setStatus("Company profile error: " + e.message))
  );
  qs("saveCompanyProfileBtn")?.addEventListener("click", () =>
    saveCompanyProfile().catch((e) => setStatus("Company profile error: " + e.message))
  );
  qs("loadAdminUsersBtn")?.addEventListener("click", () =>
    loadAdminUsers().catch((e) => setStatus("Admin users error: " + e.message))
  );
  qs("mdmLoadBtn")?.addEventListener("click", () =>
    loadMasterDataGovernance().catch((e) => setStatus("Master data error: " + e.message))
  );
  qs("mdmDeptSaveBtn")?.addEventListener("click", () =>
    saveMdmDepartment().catch((e) => setStatus("Department error: " + e.message))
  );
  qs("mdmCcSaveBtn")?.addEventListener("click", () =>
    saveMdmCostCenter().catch((e) => setStatus("Cost center error: " + e.message))
  );
  qs("mdmSupSaveBtn")?.addEventListener("click", () =>
    saveMdmSupplier().catch((e) => setStatus("Supplier error: " + e.message))
  );
  qs("mdmPolicySaveBtn")?.addEventListener("click", () =>
    saveMdmPolicies().catch((e) => setStatus("Policy error: " + e.message))
  );
  qs("adminArtisanPresetBtn")?.addEventListener("click", applyAdminArtisanPreset);
  qs("saveAdminUserBtn")?.addEventListener("click", () =>
    saveAdminUser().catch((e) => setStatus("Save user error: " + e.message))
  );
  qs("chPwdSubmit")?.addEventListener("click", () =>
    submitChangePassword().catch((e) => setStatus("Password error: " + e.message))
  );
  qs("loadSmtpSettingsBtn")?.addEventListener("click", () =>
    loadSmtpSettings().catch((e) => setStatus("SMTP load error: " + e.message))
  );
  qs("saveSmtpSettingsBtn")?.addEventListener("click", () =>
    saveSmtpSettings().catch((e) => setStatus("SMTP save error: " + e.message))
  );
  qs("testSmtpSettingsBtn")?.addEventListener("click", () =>
    testSmtpSettings().catch((e) => setStatus("SMTP test error: " + e.message))
  );
  qs("loadPushNotifyBtn")?.addEventListener("click", () =>
    loadPushNotificationSettings().catch((e) => setStatus("Push load error: " + e.message))
  );
  qs("sendPushNotifyTestBtn")?.addEventListener("click", () =>
    sendPushNotificationTest().catch((e) => setStatus("Push test error: " + e.message))
  );
  qs("sendPushNotifyBtn")?.addEventListener("click", () =>
    sendPushNotificationManual().catch((e) => setStatus("Push send error: " + e.message))
  );
  qs("loadPdfReportSettingsBtn")?.addEventListener("click", () =>
    loadPdfReportSettings().catch((e) => setStatus("PDF site load error: " + e.message))
  );
  qs("savePdfReportSettingsBtn")?.addEventListener("click", () =>
    savePdfReportSettings().catch((e) => setStatus("PDF site save error: " + e.message))
  );
  qs("uploadPdfReportLogoBtn")?.addEventListener("click", () =>
    uploadPdfReportLogo().catch((e) => setStatus("PDF logo upload error: " + e.message))
  );
  qs("removePdfReportLogoBtn")?.addEventListener("click", () =>
    removePdfReportLogo().catch((e) => setStatus("PDF logo remove error: " + e.message))
  );
  qs("pdfReportSiteCode")?.addEventListener("change", () => onPdfReportSiteCodeChange());
  qs("pdfReportCompanyCode")?.addEventListener("change", () => onPdfReportCompanyCodeChange());
  qs("sendSubscriptionNowBtn")?.addEventListener("click", () =>
    sendSubscriptionNowFromAdmin().catch((e) => setStatus("Subscription send error: " + e.message))
  );
  qs("refreshBackupsBtn")?.addEventListener("click", () =>
    loadBackupFiles().catch((e) => setStatus("Backups load error: " + e.message))
  );
  qs("createBackupNowBtn")?.addEventListener("click", () =>
    createBackupNow().catch((e) => setStatus("Backup create error: " + e.message))
  );
  qs("previewBackupRestoreBtn")?.addEventListener("click", () =>
    previewBackupRestore().catch((e) => setStatus("Backup preview error: " + e.message))
  );
  qs("stageBackupRestoreBtn")?.addEventListener("click", () =>
    stageBackupRestore().catch((e) => setStatus("Restore stage error: " + e.message))
  );
  qs("executeBackupRestoreBtn")?.addEventListener("click", () =>
    executeBackupRestoreNow().catch((e) => setStatus("Restore execute error: " + e.message))
  );
}

/** Start-up: Safety checklist templates, QR labels and safety PDFs. Called once from init() in init.js. */
function wireSafetyAdminControls() {
  qs("safetyTplLoadBtn")?.addEventListener("click", () =>
    loadSafetyTemplateEditor().catch((e) => setStatus("Safety template error: " + e.message))
  );
  qs("safetyTplSelect")?.addEventListener("change", () =>
    loadSafetyTemplateEditor().catch(() => {})
  );
  qs("safetyCategoryAddBtn")?.addEventListener("click", () =>
    addSafetyCategory().catch((e) => setStatus("Add category error: " + e.message))
  );
  qs("safetyCategoriesList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-safety-edit-category]");
    if (!btn) return;
    const key = String(btn.getAttribute("data-safety-edit-category") || "").trim();
    if (qs("safetyTplSelect")) qs("safetyTplSelect").value = key;
    loadSafetyTemplateEditor().catch(() => {});
  });
  qs("safetyTplSaveBtn")?.addEventListener("click", () =>
    saveSafetyTemplateEditor().catch((e) => setStatus("Safety template save error: " + e.message))
  );
  qs("safetyTplAddRowBtn")?.addEventListener("click", () => {
    safetyTplItems.push({ key: `item_${safetyTplItems.length + 1}`, label: "New checklist item" });
    renderSafetyTemplateEditor(safetyTplItems);
  });
  qs("safetyTplItems")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-safety-tpl-remove]");
    if (!btn) return;
    const idx = Number(btn.getAttribute("data-safety-tpl-remove"));
    safetyTplItems.splice(idx, 1);
    renderSafetyTemplateEditor(safetyTplItems);
  });
  qs("safetyItemAddBtn")?.addEventListener("click", () =>
    addSafetyEquipmentItem().catch((e) => setStatus("Add safety item error: " + e.message))
  );
  qs("safetyItemsList")?.addEventListener("click", (e) => {
    const qrBtn = e.target.closest("button[data-safety-use-qr]");
    if (qrBtn) {
      const code = String(qrBtn.getAttribute("data-safety-use-qr") || "");
      if (qs("safetyQrItemCode")) qs("safetyQrItemCode").value = code;
      generateSafetyQr().catch((err) => setStatus("QR error: " + err.message));
      return;
    }
    const inspBtn = e.target.closest("button[data-safety-open-insp]");
    if (inspBtn) {
      const code = String(inspBtn.getAttribute("data-safety-open-insp") || "");
      if (code) window.open(`./safety-inspection.html?item_code=${encodeURIComponent(code)}`, "_blank");
      return;
    }
    const pdfBtn = e.target.closest("button[data-safety-item-pdf]");
    if (pdfBtn) {
      openSafetyItemInspectionPdf(pdfBtn.getAttribute("data-safety-item-pdf"))
        .catch((err) => setStatus("Safety PDF error: " + err.message));
      return;
    }
    const rmBtn = e.target.closest("button[data-safety-remove-item]");
    if (rmBtn) {
      removeSafetyEquipmentItem(Number(rmBtn.getAttribute("data-safety-remove-item") || 0))
        .catch((err) => setStatus("Remove error: " + err.message));
    }
  });
  qs("safetyItemsList")?.addEventListener("change", (e) => {
    const chk = e.target.closest("input[data-safety-report-select]");
    if (!chk) return;
    const code = String(chk.getAttribute("data-safety-report-select") || "").trim().toUpperCase();
    if (!code) return;
    if (chk.checked) safetyReportSelectedCodes.add(code);
    else safetyReportSelectedCodes.delete(code);
  });
  qs("safetyQrGenerate")?.addEventListener("click", () =>
    generateSafetyQr().catch((e) => setStatus("Safety QR error: " + e.message))
  );
  qs("safetyQrPrint")?.addEventListener("click", printSafetyQr);
  qs("safetyQrPrintSheet")?.addEventListener("click", () =>
    printAllSafetyQrSheet().catch((e) => setStatus("Safety QR sheet error: " + e.message))
  );
  qs("safetyQrPreset")?.addEventListener("change", applySafetyQrSheetPreset);
  ["safetyQrCols", "safetyQrSizeMm", "safetyQrCellMm", "safetyQrGapMm"].forEach((id) => {
    qs(id)?.addEventListener("input", () => {
      const preset = qs("safetyQrPreset");
      if (preset && preset.value !== "custom") preset.value = "custom";
    });
  });
  qs("safetyPdfRegisterBtn")?.addEventListener("click", () =>
    openSafetyRegisterPdf(false).catch((e) => setStatus("Safety PDF error: " + e.message))
  );
  qs("safetyPdfBlankBtn")?.addEventListener("click", () =>
    openSafetyRegisterPdf(true).catch((e) => setStatus("Safety PDF error: " + e.message))
  );
  qs("safetyInspectionReportAllBtn")?.addEventListener("click", () =>
    openSafetyInspectionReportPdf(false).catch((e) => setStatus("Safety report error: " + e.message))
  );
  qs("safetyInspectionReportSelectedBtn")?.addEventListener("click", () =>
    openSafetyInspectionReportPdf(true).catch((e) => setStatus("Safety report error: " + e.message))
  );
  initSafetyAdminPanel().catch(() => {});
  initOfflineQueueAdminPanel();
  initTelematicsAdminPanel().catch(() => {});
}

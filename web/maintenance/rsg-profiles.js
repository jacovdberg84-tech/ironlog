// IRONLOG/web/maintenance/rsg-profiles.js — RSG service kit profiles.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

function parseLines(text) {
  return String(text || "")
    .split("\n")
    .map((s) => s.trim())
    .filter(Boolean);
}

function parseCsvLine(line) {
  const out = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    const nxt = line[i + 1];
    if (ch === '"' && inQuotes && nxt === '"') {
      cur += '"';
      i += 1;
      continue;
    }
    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }
    if (ch === "," && !inQuotes) {
      out.push(cur);
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur);
  return out.map((s) => String(s || "").trim());
}

function parsePackedItems(s, fallbackUnit) {
  return String(s || "")
    .split("||")
    .map((x) => String(x || "").trim())
    .filter(Boolean)
    .map((chunk) => {
      const [nameRaw, qtyRaw, unitRaw, hintRaw] = chunk.split("|").map((v) => String(v || "").trim());
      const qty = Number(qtyRaw || 0);
      return {
        name: nameRaw,
        qty: Number.isFinite(qty) ? qty : 0,
        unit: unitRaw || fallbackUnit,
        part_hint: hintRaw || nameRaw,
      };
    })
    .filter((x) => x.name && x.qty > 0);
}

function parseRsgItems(text, defaultUnit) {
  return parseLines(text).map((line) => {
    const [nameRaw, qtyRaw, unitRaw, hintRaw] = line.split("|").map((s) => String(s || "").trim());
    const qty = Number(qtyRaw || 0);
    return {
      name: nameRaw,
      qty: Number.isFinite(qty) ? qty : 0,
      unit: unitRaw || defaultUnit,
      part_hint: hintRaw || nameRaw,
    };
  }).filter((x) => x.name && x.qty > 0);
}

function itemLines(items = []) {
  return (Array.isArray(items) ? items : [])
    .map((x) => `${String(x?.name || "").trim()}|${Number(x?.qty || 0)}|${String(x?.unit || "").trim()}|${String(x?.part_hint || x?.name || "").trim()}`)
    .join("\n");
}

function slug(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function rsgProfileRow(r) {
  return `
    <tr>
      <td>${esc(r.key || "-")}</td>
      <td>${esc(r.make || "-")}</td>
      <td>${esc((r.modelContains || []).join(", ") || "-")}</td>
      <td style="text-align:right;">${Number(r.service_hours || 0)}</td>
      <td>${esc(r.title || "-")}</td>
      <td style="text-align:right;">${Array.isArray(r.oils) ? r.oils.length : 0}</td>
      <td style="text-align:right;">${Array.isArray(r.filters) ? r.filters.length : 0}</td>
      <td><button data-rsg-edit="${esc(r.key || "")}">Edit</button></td>
    </tr>
  `;
}

let rsgProfilesCache = [];
async function loadRsgProfiles() {
  const body = document.getElementById("rsgProfilesBody");
  if (!body) return;
  body.innerHTML = `<tr><td colspan="8" class="muted">Loading...</td></tr>`;
  try {
    const res = await fetch(`${API}/ironmind/rsg/profiles`);
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to load RSG profiles");
    rsgProfilesCache = Array.isArray(data.rows) ? data.rows : [];
    body.innerHTML = rsgProfilesCache.length
      ? rsgProfilesCache.map(rsgProfileRow).join("")
      : `<tr><td colspan="8" class="muted">No saved profiles yet.</td></tr>`;
  } catch (e) {
    body.innerHTML = `<tr><td colspan="8" class="message-error">${esc(e.message || String(e))}</td></tr>`;
  }
}

function fillRsgProfileForm(profileKey) {
  const p = rsgProfilesCache.find((x) => String(x.key || "") === String(profileKey || ""));
  if (!p) return;
  const set = (id, v) => {
    const el = document.getElementById(id);
    if (el) el.value = v == null ? "" : String(v);
  };
  set("rsgProfileKey", p.key || "");
  set("rsgMake", p.make || "");
  set("rsgModelMatch", Array.isArray(p.modelContains) ? p.modelContains.join(",") : "");
  set("rsgServiceHours", Number(p.service_hours || 0) || 0);
  set("rsgTitle", p.title || "");
  set("rsgTasks", parseLines((p.tasks || []).join("\n")).join("\n"));
  set("rsgChecks", parseLines((p.checks || []).join("\n")).join("\n"));
  set("rsgPostChecks", parseLines((p.post_service_checks || []).join("\n")).join("\n"));
  set("rsgSafety", parseLines((p.safety || []).join("\n")).join("\n"));
  set("rsgOils", itemLines(p.oils || []));
  set("rsgFilters", itemLines(p.filters || []));
}

async function saveRsgProfile() {
  const msg = document.getElementById("rsgProfileMsg");
  const make = String(document.getElementById("rsgMake")?.value || "").trim();
  const model_match = String(document.getElementById("rsgModelMatch")?.value || "").trim();
  const service_hours = Math.max(1, Number(document.getElementById("rsgServiceHours")?.value || 0));
  const title = String(document.getElementById("rsgTitle")?.value || "").trim();
  let profile_key = String(document.getElementById("rsgProfileKey")?.value || "").trim().toLowerCase();
  const tasks = parseLines(document.getElementById("rsgTasks")?.value || "");
  const checks = parseLines(document.getElementById("rsgChecks")?.value || "");
  const post_service_checks = parseLines(document.getElementById("rsgPostChecks")?.value || "");
  const safety = parseLines(document.getElementById("rsgSafety")?.value || "");
  const oils = parseRsgItems(document.getElementById("rsgOils")?.value || "", "L");
  const filters = parseRsgItems(document.getElementById("rsgFilters")?.value || "", "ea");
  if (!profile_key) {
    profile_key = `${slug(make)}-${slug(model_match || "model")}-${Number(service_hours || 0)}`;
    const keyEl = document.getElementById("rsgProfileKey");
    if (keyEl) keyEl.value = profile_key;
  }
  if (!profile_key || !title || !service_hours) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Profile key, service hours, and title are required.";
    }
    return;
  }
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Saving profile...";
  }
  try {
    const res = await fetch(`${API}/ironmind/rsg/profiles`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        profile_key,
        make,
        model_match,
        service_hours,
        title,
        tasks,
        checks,
        post_service_checks,
        safety,
        oils,
        filters,
      }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Failed to save profile");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `Profile saved: ${profile_key}`;
    }
    await loadRsgProfiles();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

function downloadRsgCsvTemplate() {
  return downloadProtectedXlsxFile(
    `${API}/ironmind/rsg/profiles/template.csv`,
    "IRONLOG_RSG_Profiles_Template.csv",
  ).catch((err) => alert(`Could not download RSG template: ${err.message || err}`));
}

async function importRsgProfilesCsv() {
  const msg = document.getElementById("rsgProfileMsg");
  const file = document.getElementById("rsgCsvFile")?.files?.[0];
  if (!file) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "Choose a CSV file first.";
    }
    return;
  }
  if (msg) {
    msg.className = "muted";
    msg.textContent = "Parsing CSV and importing...";
  }
  const txt = await file.text();
  const lines = String(txt || "").split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
  if (lines.length < 2) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = "CSV has no data rows.";
    }
    return;
  }
  const header = parseCsvLine(lines[0]).map((h) => h.toLowerCase());
  const idx = (name) => header.indexOf(String(name).toLowerCase());
  const get = (arr, name) => {
    const i = idx(name);
    return i >= 0 ? String(arr[i] || "").trim() : "";
  };
  const profiles = [];
  for (let i = 1; i < lines.length; i += 1) {
    const cols = parseCsvLine(lines[i]);
    const profile_key = get(cols, "profile_key");
    const service_hours = Number(get(cols, "service_hours") || 0);
    const title = get(cols, "title");
    if (!profile_key || !service_hours || !title) continue;
    profiles.push({
      profile_key,
      make: get(cols, "make"),
      model_match: get(cols, "model_match"),
      service_hours,
      title,
      tasks: parseLines(get(cols, "tasks").split("|").join("\n")),
      checks: parseLines(get(cols, "checks").split("|").join("\n")),
      post_service_checks: parseLines(get(cols, "post_service_checks").split("|").join("\n")),
      safety: parseLines(get(cols, "safety").split("|").join("\n")),
      oils: parsePackedItems(get(cols, "oils"), "L"),
      filters: parsePackedItems(get(cols, "filters"), "ea"),
    });
  }
  try {
    const res = await fetch(`${API}/ironmind/rsg/profiles/import`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ profiles }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Import failed");
    if (msg) {
      msg.className = "message-success";
      msg.textContent = `CSV import complete: ${Number(data.upserted || 0)} profile(s) upserted.`;
    }
    await loadRsgProfiles();
  } catch (e) {
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }
}

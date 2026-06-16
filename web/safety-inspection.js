(function () {
  function qs(id) {
    return document.getElementById(id);
  }
  function txt(id, value) {
    const el = qs(id);
    if (el) el.textContent = value;
  }
  function msg(text, type) {
    const box = qs("msg");
    if (!box) return;
    if (!text) {
      box.className = "";
      box.textContent = "";
      return;
    }
    box.className = `msg ${type === "ok" ? "ok" : "err"}`;
    box.textContent = text;
  }
  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    const t = await res.text();
    let data = {};
    try {
      data = t ? JSON.parse(t) : {};
    } catch {
      data = {};
    }
    if (!res.ok) throw new Error(data?.error || data?.message || t || `Request failed (${res.status})`);
    return data || {};
  }
  function getItemCodeFromUrl() {
    return String(new URL(window.location.href).searchParams.get("item_code") || "").trim().toUpperCase();
  }
  function getAssetCodeFromUrl() {
    return String(new URL(window.location.href).searchParams.get("asset_code") || "").trim().toUpperCase();
  }
  function getTemplateKeyFromUrl() {
    return String(new URL(window.location.href).searchParams.get("template_key") || "").trim().toLowerCase();
  }
  function todayYmd() {
    return new Date().toISOString().slice(0, 10);
  }
  function esc(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/"/g, "&quot;");
  }

  let currentItemCode = "";
  let currentDate = todayYmd();
  let templateItems = [];

  function renderChecklist(checklist) {
    const root = qs("checklistRoot");
    if (!root) return;
    const byKey = {};
    for (const row of Array.isArray(checklist) ? checklist : []) {
      byKey[String(row.key || "")] = row;
    }
    root.innerHTML = templateItems.map((it) => {
      const row = byKey[it.key] || {};
      let val = "";
      if (row.ok === true) val = "ok";
      else if (row.ok === false) val = "fail";
      else if (row.ok == null && row.ok !== undefined) val = "na";
      const note = String(row.note || "");
      return `
        <div class="chk-row" data-key="${esc(it.key)}">
          <div>
            <div class="chk-label">${esc(it.label)}</div>
            <input class="chk-note" type="text" data-note-key="${esc(it.key)}" placeholder="Note (required if Fail)" value="${esc(note)}" />
          </div>
          <div class="chk-opts">
            <label><input type="radio" name="chk-${esc(it.key)}" value="ok" ${val === "ok" ? "checked" : ""} /> OK</label>
            <label><input type="radio" name="chk-${esc(it.key)}" value="fail" ${val === "fail" ? "checked" : ""} /> Fail</label>
            <label><input type="radio" name="chk-${esc(it.key)}" value="na" ${val === "na" ? "checked" : ""} /> N/A</label>
          </div>
        </div>
      `;
    }).join("");
  }

  function readChecklist() {
    return templateItems.map((it) => {
      const sel = document.querySelector(`input[name="chk-${CSS.escape(it.key)}"]:checked`);
      const val = sel ? String(sel.value) : "";
      let ok = null;
      if (val === "ok") ok = true;
      else if (val === "fail") ok = false;
      const noteEl = document.querySelector(`input[data-note-key="${CSS.escape(it.key)}"]`);
      const note = String(noteEl?.value || "").trim() || null;
      return { key: it.key, label: it.label, ok, note };
    });
  }

  async function loadContext() {
    currentItemCode = getItemCodeFromUrl();
    if (!currentItemCode) {
      txt("sub", "Missing item_code in URL. Scan the equipment QR label.");
      return;
    }
    txt("sub", `Loading ${currentItemCode}...`);
    msg("");
    const q = new URLSearchParams();
    q.set("item_code", currentItemCode);
    const assetCode = getAssetCodeFromUrl();
    const templateKey = getTemplateKeyFromUrl();
    if (assetCode) q.set("asset_code", assetCode);
    if (templateKey) q.set("template_key", templateKey);
    q.set("inspection_date", currentDate);
    const data = await fetchJson(`/api/safety/inspection-context?${q.toString()}`);
    const item = data?.item || {};
    const template = data?.template || {};
    templateItems = Array.isArray(template.items) ? template.items : [];

    txt("pageTitle", String(template.title || "Safety Inspection"));
    txt("sub", "Complete all checklist items, then submit.");
    txt("itemCode", String(item.item_code || currentItemCode));
    txt("itemName", String(item.item_name || "-"));
    txt("itemLocation", String(item.location || "-"));
    txt("inspDate", String(data?.inspection_date || currentDate));

    const existing = data?.existing_inspection;
    if (qs("inspectorName")) {
      qs("inspectorName").value = String(existing?.inspector_name || "");
    }
    if (qs("notes")) {
      qs("notes").value = String(existing?.notes || "");
    }
    renderChecklist(data?.checklist || []);
    if (existing) {
      msg("Inspection already captured for this date. You can update and resubmit.", "ok");
    }
  }

  async function submitInspection() {
    msg("");
    const checklist = readChecklist();
    const body = {
      item_code: currentItemCode,
      inspection_date: currentDate,
      inspector_name: String(qs("inspectorName")?.value || "").trim(),
      notes: String(qs("notes")?.value || "").trim(),
      checklist,
    };
    const data = await fetchJson("/api/safety/inspections", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const st = String(data?.status || "pass");
    const label = st === "fail" ? "Failed — action required" : st === "attention" ? "Submitted with notes" : "Passed";
    msg(`Inspection saved: ${label}.`, "ok");
  }

  qs("refreshBtn")?.addEventListener("click", () => loadContext().catch((e) => msg(e.message || String(e), "err")));
  qs("saveBtn")?.addEventListener("click", () => submitInspection().catch((e) => msg(e.message || String(e), "err")));

  loadContext().catch((e) => {
    txt("sub", "Load failed.");
    msg(e.message || String(e), "err");
  });
})();

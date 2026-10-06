// IRONLOG/web/app/lube-model.js — Monthly lube model card (Stock Control → Reports).
// Part of the main app; index.html loads these files in order and they share one global scope.
//
// The site's lube costing workbook is uploaded once; for a month IronLog fills
// its input cells (issues per day, deliveries, opening stock) and downloads it.

let lmOilTypes = [];

function lmHeaders(extra = {}) {
  const h = {
    ...extra,
    "x-user-name": getSessionUser(),
    "x-user-role": getSessionRole(),
    "x-user-roles": getSessionRoles().join(","),
    "x-site-code": getSessionSite(),
  };
  const tok = getAuthToken();
  if (tok) h.Authorization = `Bearer ${tok}`;
  return h;
}

function lmMsg(html, kind = "") {
  const el = qs("lmMsg");
  if (!el) return;
  el.className = `dlv-msg ${kind}`;
  el.innerHTML = html;
}

function lmMonth() {
  return String(qs("lmMonth")?.value || "").trim();
}

function lmDefaultMonth() {
  // The month just closed: the model is submitted after month end.
  const d = new Date();
  d.setDate(1);
  d.setMonth(d.getMonth() - 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

async function loadLubeModel() {
  const data = await fetchJson(`${API}/api/stock/lube-model`);
  lmOilTypes = Array.isArray(data.oil_types) ? data.oil_types : [];
  const t = data.template;
  if (qs("lmTemplate")) {
    qs("lmTemplate").innerHTML = t
      ? `Template: <b>${escapeHtml(t.file_name)}</b> <span class="muted small">uploaded ${escapeHtml(String(t.uploaded_at || "").slice(0, 16))}${t.uploaded_by ? ` by ${escapeHtml(t.uploaded_by)}` : ""}</span>`
      : `<span class="muted">No template yet: upload your lube model workbook once (again when a new version comes).</span>`;
  }
  renderLubeModelParts(Array.isArray(data.parts) ? data.parts : []);
}

function renderLubeModelParts(parts) {
  const host = qs("lmParts");
  if (!host) return;
  if (!parts.length) {
    host.innerHTML = `<div class="muted small">No lube stock items yet.</div>`;
    return;
  }
  const opts = (cur) => `<option value="">— choose —</option>${lmOilTypes.map((t) => `<option ${t === cur ? "selected" : ""}>${escapeHtml(t)}</option>`).join("")}`;
  host.innerHTML = `<table class="dlv-table">
    <thead><tr><th>Item</th><th>Oil type in the model</th><th class="dlv-num">Litres/kg per unit</th><th></th></tr></thead>
    <tbody>${parts.map((p) => `
      <tr data-lm-part="${escapeHtml(p.part_code)}" class="${p.oil_type ? "" : "lm-missing"}">
        <td><b>${escapeHtml(p.part_code)}</b><br><small class="muted">${escapeHtml(p.part_name || "")}</small></td>
        <td><select data-lm-type>${opts(p.oil_type)}</select>${p.oil_type && !p.oil_type_saved ? `<br><small class="muted">suggested</small>` : ""}</td>
        <td class="dlv-num"><input data-lm-pack type="number" min="0" step="any" value="${escapeHtml(String(p.pack_size ?? ""))}" style="max-width:90px" />${p.pack_size_found ? "" : `<br><small class="muted">size not in name</small>`}</td>
        <td><button type="button" class="btn btn-secondary btn-sm" data-lm-save>Save</button></td>
      </tr>`).join("")}</tbody></table>`;
}

async function previewLubeModel() {
  const month = lmMonth();
  if (!month) return lmMsg("Choose the month.", "bad");
  lmMsg("Checking…");
  const d = await fetchJson(`${API}/api/stock/lube-model/preview?month=${encodeURIComponent(month)}`);
  const warn = [];
  if (!d.template) warn.push("Upload the template before downloading.");
  if (d.unmapped?.length) warn.push(`${d.unmapped.length} item${d.unmapped.length === 1 ? " has" : "s have"} no oil type and will be left out: ${d.unmapped.map((u) => escapeHtml(u.part_code)).join(", ")}. Set it under "Lube items" below.`);
  if (d.too_many_days) warn.push("More issue days than the model's 50 input sheets.");
  if (!(d.opening || []).length) warn.push("No lube stock in IronLog before this month: the opening stock in the template is left as it is.");
  lmMsg(warn.length ? warn.join("<br>") : "Ready to download.", warn.length ? "warn" : "ok");
  if (d.unmapped?.length && qs("lmPartsBox")) qs("lmPartsBox").open = true;
  const sum = (list) => list.reduce((s, x) => s + Number(x.qty || 0), 0);
  qs("lmSummary").innerHTML = `
    <div class="lm-stats">
      <div><b>${d.issue_days}</b><span>issue days</span></div>
      <div><b>${d.issue_lines}</b><span>issue lines</span></div>
      <div><b>${Number(d.issue_litres || 0).toLocaleString()}</b><span>litres / kg issued</span></div>
      <div><b>${(d.deliveries || []).length}</b><span>delivery lines</span></div>
      <div><b>${Math.round(sum(d.opening || [])).toLocaleString()}</b><span>opening stock (L)</span></div>
    </div>
    ${(d.issues_by_type || []).length ? `<table class="dlv-table"><thead><tr><th>Oil type</th><th class="dlv-num">Issued</th><th class="dlv-num">Opening stock</th></tr></thead><tbody>${d.issues_by_type.map((t) => `<tr><td>${escapeHtml(t.oil_type)}</td><td class="dlv-num">${Number(t.qty).toLocaleString()}</td><td class="dlv-num">${Number((d.opening || []).find((o) => o.oil_type === t.oil_type)?.qty || 0).toLocaleString()}</td></tr>`).join("")}</tbody></table>` : `<p class="muted">No lube issues in ${escapeHtml(d.month)}.</p>`}
    ${(d.deliveries || []).length ? `<p class="muted small">Deliveries: ${d.deliveries.map((x) => `${escapeHtml(x.invoice || "no invoice")} ${escapeHtml(x.description)} ×${x.items}`).join(" · ")}</p>` : ""}`;
}

async function downloadLubeModel() {
  const month = lmMonth();
  if (!month) return lmMsg("Choose the month.", "bad");
  const btn = qs("lmDownload");
  if (btn) btn.disabled = true;
  lmMsg("Filling the model… this takes a few seconds.");
  try {
    const res = await fetch(`${API}/api/stock/lube-model/download?month=${encodeURIComponent(month)}`, { headers: lmHeaders() });
    if (!res.ok) {
      let msg = `Could not fill the model (${res.status})`;
      try { msg = (await res.json()).error || msg; } catch (_) { /* not JSON */ }
      throw new Error(msg);
    }
    const name = /filename="([^"]+)"/.exec(res.headers.get("Content-Disposition") || "")?.[1] || `LUBE_MODEL_${month}.xlsm`;
    const url = URL.createObjectURL(await res.blob());
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    lmMsg(`Downloaded <b>${escapeHtml(name)}</b>. Open it in Excel (enable content) and check the summary before submitting.`, "ok");
  } catch (e) {
    lmMsg(escapeHtml(e.message || String(e)), "bad");
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function uploadLubeModelTemplate(file) {
  if (!file) return;
  lmMsg("Uploading template…");
  const fd = new FormData();
  fd.append("file", file, file.name);
  const res = await fetch(`${API}/api/stock/lube-model/template`, { method: "POST", headers: lmHeaders(), body: fd });
  let data = {};
  try { data = await res.json(); } catch (_) { data = {}; }
  if (!res.ok) throw new Error(data.error || `Upload failed (${res.status})`);
  lmMsg(`Template saved: <b>${escapeHtml(data.template?.file_name || file.name)}</b>.`, "ok");
  await loadLubeModel();
}

/** Start-up: Monthly lube model card. Called once from init() in init.js. */
function wireLubeModel() {
  const card = qs("lubeModelCard");
  if (!card) return;
  if (qs("lmMonth") && !qs("lmMonth").value) qs("lmMonth").value = lmDefaultMonth();
  let loaded = false;
  const ensureLoaded = () => {
    if (loaded) return;
    loaded = true;
    loadLubeModel().catch((e) => lmMsg(escapeHtml(e.message || String(e)), "bad"));
  };
  qs("stockTabs")?.addEventListener("click", (e) => {
    if (e.target.closest('[data-stock-tab-btn="reports"]')) ensureLoaded();
  });
  if (!card.hidden) ensureLoaded();
  qs("lmPreview")?.addEventListener("click", () => previewLubeModel().catch((e) => lmMsg(escapeHtml(e.message || String(e)), "bad")));
  qs("lmDownload")?.addEventListener("click", () => downloadLubeModel());
  qs("lmFile")?.addEventListener("change", (e) => {
    uploadLubeModelTemplate(e.target.files?.[0]).catch((err) => lmMsg(escapeHtml(err.message || String(err)), "bad"));
    e.target.value = "";
  });
  qs("lmParts")?.addEventListener("click", async (e) => {
    const btn = e.target.closest("[data-lm-save]");
    if (!btn) return;
    const row = btn.closest("[data-lm-part]");
    btn.disabled = true;
    try {
      await fetchJson(`${API}/api/stock/lube-model/parts/${encodeURIComponent(row.dataset.lmPart)}`, {
        method: "PUT",
        body: JSON.stringify({ oil_type: row.querySelector("[data-lm-type]").value, pack_size: row.querySelector("[data-lm-pack]").value }),
      });
      row.classList.toggle("lm-missing", !row.querySelector("[data-lm-type]").value);
      btn.textContent = "Saved";
      setTimeout(() => { btn.textContent = "Save"; }, 1500);
    } catch (err) {
      lmMsg(escapeHtml(err.message || String(err)), "bad");
    } finally {
      btn.disabled = false;
    }
  });
}

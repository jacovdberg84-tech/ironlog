// IRONLOG/web/app/licence-issuer.js — Admin → Issue licences (the IronLog owner only).
// Part of index.html; the page loads these files in order and they share one global scope.
// Packages, machine bands and optional prices are edited here; issuing signs a
// licence key with this server's signing key (api/utils/licenceIssuer.js).

let liCatalogue = null;

const liApi = (path, opts = {}) => fetchJson(`${API}/api/licence/issuer${path}`, opts);
const liMoney = (v, ccy) => (v == null ? "—" : `${ccy || "USD"} ${Number(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);

async function loadLicenceIssuer() {
  const card = qs("licenceIssuerCard");
  if (!card) return;
  const st = await liApi("/status").catch(() => ({ allowed: false }));
  card.classList.toggle("hidden", !st.allowed);
  if (!st.allowed) return;
  const keyBox = qs("liKeyState");
  if (!st.has_signing_key) {
    keyBox.innerHTML = `<p>This server has no signing key yet. Set it up once; it stays on this server and is never shown again.</p>
      <button type="button" id="liSetupBtn" class="btn-primary">Set up the signing key</button>`;
    qs("liIssueSection").classList.add("hidden");
    return;
  }
  keyBox.innerHTML = `
    <p class="muted small">Licences are signed on this server. ${st.built_in ? "Every IronLog server trusts this key." : "Customer servers trust it once the public key below is built into IronLog: send it to Claude to add."}</p>
    <label class="muted small" for="liPublicKey">Public key (safe to share)</label>
    <input id="liPublicKey" class="w-full" readonly value="${escapeHtml(st.public_key)}" />
    <div class="row" style="gap:8px; margin-top:8px; flex-wrap:wrap;">
      <button type="button" id="liCopyPubBtn" class="btn btn-secondary btn-sm">Copy public key</button>
      <button type="button" id="liDownloadKeyBtn" class="btn btn-secondary btn-sm" title="Keep this file offline (USB stick in a safe). Anyone with it can issue IronLog licences.">Download a copy of the signing key</button>
    </div>`;
  qs("liIssueSection").classList.remove("hidden");
  liCatalogue = await liApi("/catalogue");
  renderLicenceCatalogue();
}

function liModuleChecks(name, selected = [], exclude = []) {
  return Object.entries(liCatalogue.modules)
    .filter(([k]) => !exclude.includes(k))
    .map(([k, label]) => `<label class="checkbox-label li-mod"><input type="checkbox" name="${name}" value="${k}"${selected.includes(k) ? " checked" : ""} /> <span>${escapeHtml(label)}</span></label>`)
    .join("");
}

function renderLicenceCatalogue() {
  const c = liCatalogue;
  qs("liPackages").innerHTML = c.packages.map((p) => `
    <div class="li-package" data-package-id="${p.id}">
      <div class="form-row">
        <div class="form-field"><label>Package</label><input class="w-full li-pkg-name" value="${escapeHtml(p.name)}" /></div>
        <div class="form-field"><label>Description</label><input class="w-full li-pkg-desc" value="${escapeHtml(p.description)}" /></div>
      </div>
      <div class="li-mods">${liModuleChecks(`pkg-${p.id}`, p.modules)}</div>
      <button type="button" class="btn btn-secondary btn-sm li-pkg-save">Save package</button>
    </div>`).join("");
  qs("liBands").innerHTML = c.bands.map((b) => `
    <div class="form-row li-band" data-band-id="${b.id}">
      <div class="form-field"><label>Band</label><input class="w-full li-band-label" value="${escapeHtml(b.label)}" /></div>
      <div class="form-field"><label>Machines (blank = unlimited)</label><input class="w-full li-band-max" type="number" min="1" value="${b.max_assets ?? ""}" /></div>
      <div class="form-field" style="align-self:end;"><button type="button" class="btn btn-secondary btn-sm li-band-save">Save band</button></div>
    </div>`).join("");
  const price = (pid, bid) => c.prices.find((p) => p.package_id === pid && p.band_id === bid);
  qs("liPrices").innerHTML = `<table class="backup-runs-table"><thead><tr><th>Monthly price per site</th>${c.bands.map((b) => `<th>${escapeHtml(b.label)}</th>`).join("")}</tr></thead><tbody>
    ${c.packages.map((p) => `<tr><td>${escapeHtml(p.name)}</td>${c.bands.map((b) => `<td><input class="li-price" type="number" min="0" step="0.01" data-p="${p.id}" data-b="${b.id}" value="${price(p.id, b.id)?.monthly_price ?? ""}" placeholder="—" /></td>`).join("")}</tr>`).join("")}
    </tbody></table>`;
  if (!qs("liSites").children.length) liAddSiteRow();
  qs("liIssuedBody").innerHTML = (c.issued || []).map((r) => `<tr>
      <td>${escapeHtml(r.licence_id)}</td><td>${escapeHtml(r.licensee)}</td><td>${escapeHtml(r.expires_at)}</td>
      <td>${escapeHtml(Object.keys(r.payload?.sites || {}).join(", "))}</td>
      <td>${r.quote?.total ? escapeHtml(liMoney(r.quote.total.monthly, r.quote.total.currency)) : "—"}</td>
      <td>${escapeHtml(String(r.issued_at || "").slice(0, 16))}</td>
      <td><button type="button" class="btn btn-secondary btn-sm li-copy-issued" data-id="${r.id}">Copy key</button></td></tr>`).join("")
    || `<tr><td colspan="7" class="muted">No licences issued yet.</td></tr>`;
}

function liAddSiteRow() {
  const c = liCatalogue;
  const row = document.createElement("div");
  row.className = "li-site";
  row.innerHTML = `
    <div class="form-row">
      <div class="form-field"><label>Site code (as in the customer's IronLog, or * for every site)</label><input class="w-full li-site-code" placeholder="main" /></div>
      <div class="form-field"><label>Site name</label><input class="w-full li-site-name" placeholder="e.g. North pit" /></div>
      <div class="form-field"><label>Package</label><select class="w-full li-site-pkg">${c.packages.map((p) => `<option value="${p.id}">${escapeHtml(p.name)}</option>`).join("")}</select></div>
      <div class="form-field"><label>Machine band</label><select class="w-full li-site-band">${c.bands.map((b) => `<option value="${b.id}">${escapeHtml(b.label)}</option>`).join("")}</select></div>
    </div>
    <details><summary class="muted small">Extra modules on top of the package</summary><div class="li-mods">${liModuleChecks("extra", [])}</div></details>
    <button type="button" class="link-btn li-site-remove">Remove site</button>`;
  qs("liSites").appendChild(row);
}

function wireLicenceIssuer() {
  const card = qs("licenceIssuerCard");
  if (!card) return;
  qs("refreshLicenceIssuerBtn")?.addEventListener("click", () => loadLicenceIssuer().catch((e) => setStatus("Licence issuer error: " + e.message)));
  card.addEventListener("click", async (ev) => {
    const t = ev.target.closest("button");
    if (!t) return;
    const out = qs("liResult");
    const say = (m) => { if (out) out.textContent = m; };
    try {
      if (t.id === "liSetupBtn") {
        if (!confirm("Set up the licence signing key on this server? It is made once and never replaced.")) return;
        await liApi("/setup", { method: "POST", body: "{}" });
        await loadLicenceIssuer();
        say("Signing key set up. Copy the public key and send it to Claude to build into IronLog.");
      } else if (t.id === "liCopyPubBtn") {
        await navigator.clipboard.writeText(qs("liPublicKey").value);
        say("Public key copied.");
      } else if (t.id === "liDownloadKeyBtn") {
        if (!confirm("Download the signing key? Keep it offline: anyone with this file can issue IronLog licences.")) return;
        await downloadAuthedFile(`${API}/api/licence/issuer/signing-key.pem`, "ironlog-licence-signing-key.pem");
      } else if (t.classList.contains("li-pkg-save")) {
        const box = t.closest(".li-package");
        await liApi("/packages", { method: "POST", body: JSON.stringify({
          id: Number(box.dataset.packageId),
          name: box.querySelector(".li-pkg-name").value,
          description: box.querySelector(".li-pkg-desc").value,
          modules: [...box.querySelectorAll("input[type=checkbox]:checked")].map((i) => i.value),
        }) });
        say("Package saved.");
        await loadLicenceIssuer();
      } else if (t.id === "liAddPackageBtn") {
        const name = prompt("Name of the new package?");
        if (!name) return;
        await liApi("/packages", { method: "POST", body: JSON.stringify({ name, modules: [] }) });
        await loadLicenceIssuer();
      } else if (t.classList.contains("li-band-save")) {
        const row = t.closest(".li-band");
        await liApi("/bands", { method: "POST", body: JSON.stringify({ id: Number(row.dataset.bandId), label: row.querySelector(".li-band-label").value, max_assets: row.querySelector(".li-band-max").value }) });
        say("Band saved.");
        await loadLicenceIssuer();
      } else if (t.id === "liAddBandBtn") {
        const label = prompt("Name of the new band, e.g. Up to 1000 machines?");
        if (!label) return;
        const max = prompt("How many machines? Leave blank for unlimited.") || "";
        await liApi("/bands", { method: "POST", body: JSON.stringify({ label, max_assets: max }) });
        await loadLicenceIssuer();
      } else if (t.id === "liSavePricesBtn") {
        const ccy = String(qs("liCurrency").value || "USD").trim().toUpperCase();
        const prices = [...card.querySelectorAll(".li-price")].map((i) => ({ package_id: Number(i.dataset.p), band_id: Number(i.dataset.b), monthly_price: i.value, currency: ccy }));
        await liApi("/prices", { method: "PUT", body: JSON.stringify({ prices }) });
        say("Prices saved.");
        await loadLicenceIssuer();
      } else if (t.id === "liAddSiteBtn") {
        liAddSiteRow();
      } else if (t.classList.contains("li-site-remove")) {
        t.closest(".li-site").remove();
      } else if (t.id === "liIssueBtn") {
        const sites = [...card.querySelectorAll(".li-site")].map((row) => ({
          site_code: row.querySelector(".li-site-code").value,
          name: row.querySelector(".li-site-name").value,
          package_id: Number(row.querySelector(".li-site-pkg").value),
          band_id: Number(row.querySelector(".li-site-band").value),
          extra_modules: [...row.querySelectorAll("input[type=checkbox]:checked")].map((i) => i.value),
        }));
        const r = await liApi("/issue", { method: "POST", body: JSON.stringify({
          licensee: qs("liLicensee").value,
          expires_at: qs("liExpires").value,
          grace_days: qs("liGrace").value,
          sites,
        }) });
        qs("liIssuedKey").value = r.key;
        qs("liIssuedBox").classList.remove("hidden");
        const lines = r.quote.sites.map((q) => `${q.site_code}: ${q.package}, ${q.band}, ${liMoney(q.monthly_price, q.currency)} a month`);
        qs("liQuote").textContent = `${r.licence_id} for ${r.payload.licensee}, ending ${r.payload.expires_at}. ${lines.join("; ")}.${r.quote.total ? ` Total ${liMoney(r.quote.total.monthly, r.quote.total.currency)} a month.` : ""}`;
        say("Licence issued. Copy the key and send it to the customer's administrator.");
        liCatalogue = await liApi("/catalogue");
        renderLicenceCatalogue();
      } else if (t.id === "liCopyKeyBtn") {
        await navigator.clipboard.writeText(qs("liIssuedKey").value);
        say("Licence key copied.");
      } else if (t.classList.contains("li-copy-issued")) {
        const k = await liApi(`/issued/${t.dataset.id}/key`);
        await navigator.clipboard.writeText(k.licence_key);
        say(`Key for ${k.licence_id} copied.`);
      }
    } catch (e) {
      say(e.message || String(e));
    }
  });
}

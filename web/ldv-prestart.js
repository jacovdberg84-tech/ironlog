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
      box.innerHTML = "";
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
  function getAssetCodeFromUrl() {
    const url = new URL(window.location.href);
    return String(url.searchParams.get("asset_code") || "").trim().toUpperCase();
  }
  function todayYmd() {
    return new Date().toISOString().slice(0, 10);
  }
  function readChecklist() {
    return {
      brakes_ok: Boolean(qs("chkBrakes")?.checked),
      lights_ok: Boolean(qs("chkLights")?.checked),
      tyres_ok: Boolean(qs("chkTyres")?.checked),
      oil_coolant_ok: Boolean(qs("chkOilCoolant")?.checked),
      leaks_damage_ok: Boolean(qs("chkLeaks")?.checked),
      safety_items_ok: Boolean(qs("chkSafety")?.checked),
    };
  }
  function applyChecklist(checklist) {
    const byKey = {};
    for (const item of Array.isArray(checklist) ? checklist : []) {
      byKey[String(item?.key || "")] = Boolean(item?.ok);
    }
    if (qs("chkBrakes")) qs("chkBrakes").checked = byKey.brakes_ok === true;
    if (qs("chkLights")) qs("chkLights").checked = byKey.lights_ok === true;
    if (qs("chkTyres")) qs("chkTyres").checked = byKey.tyres_ok === true;
    if (qs("chkOilCoolant")) qs("chkOilCoolant").checked = byKey.oil_coolant_ok === true;
    if (qs("chkLeaks")) qs("chkLeaks").checked = byKey.leaks_damage_ok === true;
    if (qs("chkSafety")) qs("chkSafety").checked = byKey.safety_items_ok === true;
  }

  let currentAssetCode = "";
  let currentDate = todayYmd();
  let previousKm = null;

  async function loadContext() {
    currentAssetCode = getAssetCodeFromUrl();
    if (!currentAssetCode) {
      txt("sub", "Missing asset_code in URL.");
      return;
    }
    txt("sub", `Loading ${currentAssetCode}...`);
    msg("");
    const q = new URLSearchParams();
    q.set("asset_code", currentAssetCode);
    q.set("check_date", currentDate);
    const data = await fetchJson(`/api/maintenance/vehicle-ldv-checks/prestart-context?${q.toString()}`);
    const asset = data?.asset || {};
    previousKm = data?.previous_odometer_km == null ? null : Number(data.previous_odometer_km);

    txt("sub", "Complete pre-start before operating vehicle.");
    txt("assetCode", String(asset.asset_code || currentAssetCode));
    txt("assetName", String(asset.asset_name || "-"));
    txt("checkDate", String(data?.check_date || currentDate));
    txt("prevKm", previousKm == null ? "-" : `${previousKm.toFixed(1)} km`);

    const existing = data?.existing_prestart || null;
    if (existing) {
      if (qs("odometerKm") && existing.odometer_km != null) qs("odometerKm").value = String(existing.odometer_km);
      if (qs("inspectorName")) qs("inspectorName").value = String(existing.inspector_name || "");
      if (qs("notes")) qs("notes").value = String(existing.notes || "");
      applyChecklist(existing.checklist);
      msg("Pre-start already captured for today. You can update and resubmit if needed.", "ok");
    }

    const openQrBtn = qs("openQrBtn");
    if (openQrBtn) openQrBtn.href = `./asset-qr.html?asset_code=${encodeURIComponent(currentAssetCode)}`;
  }

  async function submitPrestart() {
    msg("");
    const odometerRaw = String(qs("odometerKm")?.value || "").trim();
    if (!odometerRaw) throw new Error("Enter current odometer KM.");
    const odometer = Number(odometerRaw);
    if (!Number.isFinite(odometer) || odometer < 0) throw new Error("Odometer must be a valid number >= 0.");
    if (previousKm != null && odometer < previousKm) {
      throw new Error(`Odometer cannot be less than previous KM (${previousKm.toFixed(1)}).`);
    }

    const checklist = readChecklist();
    const allChecked = Object.values(checklist).every(Boolean);
    if (!allChecked) throw new Error("Complete all pre-start checks before submitting.");

    const body = {
      asset_code: currentAssetCode,
      check_date: currentDate,
      odometer_km: odometer,
      inspector_name: String(qs("inspectorName")?.value || "").trim(),
      notes: String(qs("notes")?.value || "").trim(),
      checklist,
    };
    const data = await fetchJson("/api/maintenance/vehicle-ldv-checks/prestart", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const savedKm = data?.odometer_km == null ? odometer : Number(data.odometer_km);
    previousKm = Number.isFinite(savedKm) ? savedKm : previousKm;
    txt("prevKm", previousKm == null ? "-" : `${previousKm.toFixed(1)} km`);
    msg("Pre-start submitted successfully. KM saved to IRONLOG.", "ok");
  }

  qs("refreshBtn")?.addEventListener("click", () => {
    loadContext().catch((e) => msg(String(e.message || e), "err"));
  });
  qs("saveBtn")?.addEventListener("click", () => {
    submitPrestart().catch((e) => msg(String(e.message || e), "err"));
  });

  loadContext().catch((e) => msg(String(e.message || e), "err"));
})();

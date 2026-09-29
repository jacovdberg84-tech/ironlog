// IRONLOG/web/app/uploads.js — CSV uploads, FAMS fuel import and templates.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function doUpload() {
  const endpointEl = qs("uploadEndpoint");
  const fileEl = qs("uploadFile");
  const resultEl = qs("uploadResult");
  if (!endpointEl || !fileEl || !resultEl) return;

  const endpoint = endpointEl.value;
  const file = fileEl.files[0];
  if (!file) return alert("Choose a CSV file first.");

  const fd = new FormData();
  fd.append("file", file);

  setStatus("Uploading CSV...");
  try {
    const res = await fetchJson(`${API}${endpoint}`, { method: "POST", body: fd });
    resultEl.textContent = JSON.stringify(res, null, 2);
    setStatus("Upload complete.");
    await loadDashboard().catch(() => {});
  } catch (e) {
    resultEl.textContent = String(e.message || e);
    setStatus("Upload failed.");
  }
}

async function importFamsFuelFile() {
  const fileEl = qs("fuelFamsFile");
  const modeEl = qs("fuelFamsConflictMode");
  const resultEl = qs("fuelFamsResult");
  if (!fileEl || !resultEl) return;

  const file = fileEl.files[0];
  if (!file) return alert("Choose a FAMS CSV file first.");

  const fd = new FormData();
  fd.append("file", file);
  const conflictMode = String(modeEl?.value || "skip").trim().toLowerCase();
  const mode = ["skip", "overwrite"].includes(conflictMode) ? conflictMode : "skip";

  setStatus("Importing FAMS fuel file...");
  resultEl.textContent = "";
  try {
    const res = await fetchJson(`${API}/api/upload/fuel?on_conflict=${encodeURIComponent(mode)}`, { method: "POST", body: fd });
    resultEl.textContent = JSON.stringify(res, null, 2);
    setStatus("FAMS fuel import complete.");
    await loadDashboard().catch(() => {});
    loadFamsFuelStatus().catch(() => {});
  } catch (e) {
    resultEl.textContent = String(e.message || e);
    setStatus("FAMS fuel import failed.");
  }
}

function renderFamsFuelStatus(data) {
  const status = String(data?.status || (data?.enabled ? "idle" : "disabled"));
  const labelMap = {
    connected: "Connected",
    error: "Error",
    disabled: "Disabled",
    idle: "Idle",
    syncing: "Syncing…",
  };
  setText("fuelFamsStatusLabel", labelMap[status] || status);
  setText("fuelFamsLastSuccess", data?.last_success_at ? String(data.last_success_at).replace("T", " ").slice(0, 19) : "—");
  setText("fuelFamsLastAttempt", data?.last_attempt_at ? String(data.last_attempt_at).replace("T", " ").slice(0, 19) : "—");
  setText("fuelFamsReceived", data?.last_received != null ? String(data.last_received) : "—");
  setText("fuelFamsImported", data?.last_imported != null ? String(data.last_imported) : "—");
  setText("fuelFamsSkipped", data?.last_skipped != null ? String(data.last_skipped) : "—");
  const unmatched = data?.unmatched_open != null ? data.unmatched_open : data?.last_unmatched;
  setText("fuelFamsUnmatched", unmatched != null ? String(unmatched) : "—");
  const msg = qs("fuelFamsStatusMsg");
  if (msg) {
    if (!data?.enabled) {
      msg.textContent = "Auto sync is off. Set FAMS_ENABLED=true and FAMS_AUTH in api/.env, then restart the API.";
    } else if (!data?.configured) {
      msg.textContent = "FAMS_AUTH is missing on the server.";
    } else if (data?.last_error) {
      msg.textContent = `Last error: ${data.last_error}`;
    } else if (data?.last_range_start && data?.last_range_end) {
      msg.textContent = `Last range ${data.last_range_start} → ${data.last_range_end}. Fuel only — hours stay on QR pre-start.`;
    } else {
      msg.textContent = data?.note || "";
    }
  }
}

async function loadFamsFuelStatus() {
  try {
    const data = await fetchJson(`${API}/api/dashboard/fuel/fams/status`);
    renderFamsFuelStatus(data);
    return data;
  } catch (e) {
    setText("fuelFamsStatusLabel", "Error");
    const msg = qs("fuelFamsStatusMsg");
    if (msg) msg.textContent = e.message || String(e);
    throw e;
  }
}

function initFamsFuelCatchupDates() {
  const today = todayLocalYmd();
  const monthStart = `${today.slice(0, 7)}-01`;
  const from = qs("fuelFamsSyncFromDate");
  const to = qs("fuelFamsSyncToDate");
  if (from && !from.value) from.value = monthStart;
  if (to && !to.value) to.value = today;
}

function selectedFamsFuelSyncRange() {
  const startDate = String(qs("fuelFamsSyncFromDate")?.value || "").trim();
  const endDate = String(qs("fuelFamsSyncToDate")?.value || "").trim();
  if (!startDate || !endDate) {
    alert("Select both a Catch-up From Date and Catch-up To Date.");
    return null;
  }
  if (startDate > endDate) {
    alert("Catch-up From Date cannot be after Catch-up To Date.");
    return null;
  }
  return { startDate, endDate };
}

async function syncFamsFuelNow({ startDate = "", endDate = "" } = {}) {
  const resultEl = qs("fuelFamsResult");
  const selectedRange = startDate && endDate ? ` for ${startDate} to ${endDate}` : "";
  setStatus(`Syncing FAMS fuel${selectedRange}…`);
  if (resultEl) resultEl.textContent = "Syncing…";
  try {
    const res = await fetchJson(`${API}/api/dashboard/fuel/fams/sync`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({ force: true, start_date: startDate || undefined, end_date: endDate || undefined }),
    });
    if (resultEl) resultEl.textContent = JSON.stringify(res, null, 2);
    if (res?.status) renderFamsFuelStatus(res.status);
    else await loadFamsFuelStatus().catch(() => {});
    if (res?.ok) {
      const linkedLegacy = Number(res.legacy_linked || 0);
      setStatus(
        `FAMS sync done: ${Number(res.imported || 0)} imported, ${linkedLegacy} legacy row(s) linked, ${Number(res.skipped || 0)} skipped, ${Number(res.unmatched || 0)} unmatched.`
      );
      await loadDashboard().catch(() => {});
    } else {
      setStatus(`FAMS sync failed: ${res?.error || res?.reason || "unknown"}`);
    }
  } catch (e) {
    if (resultEl) resultEl.textContent = String(e.message || e);
    setStatus("FAMS sync failed.");
    loadFamsFuelStatus().catch(() => {});
  }
}

async function repairFuelMeterChain() {
  const resultEl = qs("fuelFamsResult");
  const assetCode = (qs("fuelRepairAssetCode")?.value || "").trim();
  if (resultEl) resultEl.textContent = "";
  setStatus("Repairing meter chain...");
  try {
    const body = assetCode ? { asset_code: assetCode } : {};
    const res = await fetchJson(`${API}/api/dashboard/fuel/repair-meter-chain`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify(body),
    });
    if (resultEl) resultEl.textContent = JSON.stringify(res, null, 2);
    setStatus(`Meter chain repair complete. Rows repaired: ${Number(res?.repaired_rows || 0)}`);
    await loadDashboard().catch(() => {});
  } catch (e) {
    if (resultEl) resultEl.textContent = String(e.message || e);
    setStatus("Meter chain repair failed.");
  }
}

async function clearFuelFromDate() {
  const resultEl = qs("fuelFamsResult");
  const fromDate = (qs("fuelClearFromDate")?.value || "").trim();
  const assetCode = (qs("fuelClearAssetCode")?.value || "").trim();
  const clearDailyHours = Boolean(qs("fuelClearDailyHours")?.checked);
  if (!fromDate) return alert("Select a from date first.");

  if (resultEl) resultEl.textContent = "";
  setStatus("Checking clear impact...");
  try {
    const preview = await previewFuelClearImpact(fromDate, assetCode, clearDailyHours);
    const ok = confirm(
      `About to clear fuel from ${fromDate}` +
      `${assetCode ? ` for ${assetCode}` : " for all assets"}.\n` +
      `Fuel logs to delete: ${Number(preview?.deleted_logs || 0)}\n` +
      `Affected days: ${Number(preview?.affected_days || 0)}\n` +
      (clearDailyHours ? "Daily opening/closing/run meters on affected days will also be cleared.\n" : "") +
      "\nContinue?"
    );
    if (!ok) {
      setStatus("Fuel clear cancelled.");
      return;
    }

    setStatus("Clearing fuel data...");
    const res = await fetchJson(`${API}/api/dashboard/fuel/clear-from-date`, {
      method: "POST",
      headers: { "Content-Type": "application/json", ...authHeaders() },
      body: JSON.stringify({
        from_date: fromDate,
        ...(assetCode ? { asset_code: assetCode } : {}),
        clear_daily_hours: clearDailyHours,
      }),
    });
    if (resultEl) resultEl.textContent = JSON.stringify(res, null, 2);
    setStatus(`Fuel clear complete. Deleted logs: ${Number(res?.deleted_logs || 0)}`);
    await Promise.all([loadDashboard().catch(() => {}), loadFuelBenchmark().catch(() => {})]);
  } catch (e) {
    if (resultEl) resultEl.textContent = String(e.message || e);
    setStatus("Fuel clear failed. Check API version and permissions.");
  }
}

async function previewFuelClearImpact(fromDate, assetCode, clearDailyHours) {
  const result = await fetchJson(`${API}/api/dashboard/fuel/clear-from-date/preview`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      from_date: fromDate,
      ...(assetCode ? { asset_code: assetCode } : {}),
      clear_daily_hours: clearDailyHours,
    }),
  });
  const previewEl = qs("fuelClearPreviewResult");
  if (previewEl) previewEl.textContent = JSON.stringify(result, null, 2);
  return result;
}

async function runFuelClearPreview() {
  const fromDate = (qs("fuelClearFromDate")?.value || "").trim();
  const assetCode = (qs("fuelClearAssetCode")?.value || "").trim();
  const clearDailyHours = Boolean(qs("fuelClearDailyHours")?.checked);
  if (!fromDate) return alert("Select a from date first.");
  setStatus("Loading clear preview...");
  const preview = await previewFuelClearImpact(fromDate, assetCode, clearDailyHours);
  setStatus(
    `Preview ready. Logs: ${Number(preview?.deleted_logs || 0)}, Days: ${Number(preview?.affected_days || 0)}`
  );
}

async function editFuelMachineHours(logId, openValue, closeValue) {
  const id = Number(logId || 0);
  if (!Number.isInteger(id) || id <= 0) return;
  const openDefault = Number.isFinite(Number(openValue)) ? String(Number(openValue)) : "";
  const closeDefault = Number.isFinite(Number(closeValue)) ? String(Number(closeValue)) : "";
  const openRaw = prompt("Enter opening meter value", openDefault);
  if (openRaw == null) return;
  const closeRaw = prompt("Enter closing meter value", closeDefault);
  if (closeRaw == null) return;

  const opening = Number(String(openRaw).trim());
  const closing = Number(String(closeRaw).trim());
  if (!Number.isFinite(opening) || opening < 0) return alert("Opening meter must be >= 0.");
  if (!Number.isFinite(closing) || closing < 0) return alert("Closing meter must be >= 0.");
  if (closing < opening) return alert("Closing meter must be greater than or equal to opening meter.");

  await fetchJson(`${API}/api/dashboard/fuel/machine-hours`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      fuel_log_id: id,
      opening_meter: opening,
      closing_meter: closing,
    }),
  });
}

async function saveFuelMachineHoursInline(saveBtn) {
  const id = Number(saveBtn?.getAttribute("data-fuel-save") || 0);
  if (!Number.isInteger(id) || id <= 0) return;
  const row = saveBtn.closest("tr");
  if (!row) return;
  const openEl = row.querySelector('input[data-fuel-open-input="1"]');
  const closeEl = row.querySelector('input[data-fuel-close-input="1"]');
  const opening = Number(String(openEl?.value || "").trim());
  const closing = Number(String(closeEl?.value || "").trim());

  if (!Number.isFinite(opening) || opening < 0) return alert("Opening hours must be >= 0.");
  if (!Number.isFinite(closing) || closing < 0) return alert("Closing hours must be >= 0.");
  if (closing < opening) return alert("Closing hours must be greater than or equal to opening hours.");

  await fetchJson(`${API}/api/dashboard/fuel/machine-hours`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...authHeaders() },
    body: JSON.stringify({
      fuel_log_id: id,
      opening_meter: opening,
      closing_meter: closing,
    }),
  });
}

function downloadStoresCsvTemplate() {
  const today = new Date().toISOString().slice(0, 10);
  const lines = [
    "part_code,quantity,allocation_date,asset_code,work_order_id,issued_by,notes",
    `FLT-001,2,${today},A300AM,,Storeman A,Planned PM kit`,
    `BLT-009,1,${today},,41,Storeman B,Issued against WO 41`
  ];
  const csv = lines.join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "stores_alloc_template.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  setStatus("Stores CSV template downloaded.");
}

function downloadFuelBaselineCsvTemplate() {
  const lines = [
    "asset_code,baseline_fuel_l_per_hour",
    "A300AM,7.25",
    "A301AM,7.10",
    "E500AM,18.50"
  ];
  const csv = lines.join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "fuel_baseline_template.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  setStatus("Fuel baseline CSV template downloaded.");
}

function downloadFuelCsvTemplate() {
  const today = new Date().toISOString().slice(0, 10);
  const lines = [
    "asset_code,log_date,liters,source,meter_unit,meter_run_value,hours_run",
    `A300AM,${today},180,bowser,hours,10.0,10.0`,
    `A301AM,${today},220,bowser,hours,9.5,9.5`,
    `LDV01,${today},60,pump_1,km,480,`
  ];
  const csv = lines.join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "fuel_import_template.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  setStatus("Fuel CSV template downloaded.");
}

function saveCsvFile(csv, filename) {
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function csvCell(v) {
  const s = String(v == null ? "" : v);
  return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

async function downloadDailyHoursCsvTemplate() {
  const date = qs("date")?.value || todayLocalYmd();
  const header = "asset_code,asset_name,work_date,scheduled_hours,opening_hours,closing_hours,hours_run,is_used,operator,notes";

  setStatus("Building hours CSV template...");
  let assets = [];
  try {
    assets = await fetchJson(`${API}/api/assets?include_archived=0`);
  } catch {
    assets = [];
  }

  const active = (Array.isArray(assets) ? assets : [])
    .filter((a) => a.active !== 0 && a.active !== false)
    .sort((a, b) => String(a.asset_code || "").localeCompare(String(b.asset_code || "")));

  let lines;
  if (active.length) {
    lines = [header, ...active.map((a) => {
      const standby = a.is_standby ? 1 : 0;
      const sched = standby ? 0 : 10;
      const used = standby ? 0 : 1;
      return [
        csvCell(a.asset_code),
        csvCell(a.asset_name),
        date,
        sched,
        "",
        "",
        "",
        used,
        "",
        "",
      ].join(",");
    })];
  } else {
    // Fallback sample if the asset list could not be loaded.
    lines = [
      header,
      `A300AM,Excavator 300,${date},10,4500.0,4510.0,10.0,1,J Smith,`,
      `A301AM,Excavator 301,${date},10,3200.5,,9.5,1,,`,
      `A302AM,Dozer 302,${date},0,,,0,0,,Standby`,
    ];
  }

  saveCsvFile(lines.join("\n"), `daily_hours_template_${date}.csv`);
  setStatus(active.length
    ? `Daily hours template downloaded (${active.length} asset(s), date ${date}).`
    : "Daily hours CSV template downloaded (sample rows).");
}

async function uploadDailyHoursCsv(file) {
  if (!file) return;
  const fd = new FormData();
  fd.append("file", file);

  setStatus("Uploading hours CSV...");
  try {
    const res = await fetchJson(`${API}/api/upload/hours`, { method: "POST", body: fd });
    const imported = Number(res?.imported || 0);
    const synced = Number(res?.synced_asset_hours || 0);
    setStatus(`Hours CSV uploaded — ${imported} row(s) processed, ${synced} asset(s) synced.`);
    await loadDailyInput().catch(() => {});
  } catch (e) {
    setStatus("Hours CSV upload failed: " + (e.message || e));
    alert("Hours CSV upload failed: " + (e.message || e));
  }
}

async function downloadDailyMatrixCsvTemplate() {
  const date = qs("date")?.value || todayLocalYmd();
  setStatus("Building meter-matrix template...");
  let assets = [];
  try {
    assets = await fetchJson(`${API}/api/assets?include_archived=0`);
  } catch {
    assets = [];
  }
  const codes = (Array.isArray(assets) ? assets : [])
    .filter((a) => a.active !== 0 && a.active !== false)
    .map((a) => String(a.asset_code || "").trim())
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));

  const cols = codes.length ? codes : ["A300AM", "A301AM", "LDV01"];
  const header = ["date", ...cols].map(csvCell).join(",");
  // Two example date rows (blank cells to fill in with cumulative meter readings).
  const prev = prevDateStr(date);
  const lines = [
    header,
    [prev, ...cols.map(() => "")].join(","),
    [date, ...cols.map(() => "")].join(","),
  ];
  saveCsvFile(lines.join("\n"), `meter_matrix_template_${date}.csv`);
  setStatus(codes.length
    ? `Meter-matrix template downloaded (${codes.length} asset columns).`
    : "Meter-matrix template downloaded (sample columns).");
}

async function uploadDailyMatrixCsv(file) {
  if (!file) return;
  const scheduledRaw = Number(qs("bulkSched")?.value);
  const scheduled = Number.isFinite(scheduledRaw) && scheduledRaw > 0 && scheduledRaw <= 24 ? scheduledRaw : 10;
  const fd = new FormData();
  fd.append("file", file);

  setStatus("Uploading meter matrix...");
  try {
    const res = await fetchJson(`${API}/api/upload/hours-matrix?scheduled=${encodeURIComponent(scheduled)}`, {
      method: "POST",
      body: fd,
    });
    const rows = Number(res?.rows_written || 0);
    const assetsMatched = Number(res?.assets_matched || 0);
    const dates = Number(res?.dates || 0);
    const resets = Number(res?.meter_resets || 0);
    const unknown = Array.isArray(res?.unknown_codes) ? res.unknown_codes : [];
    let msg = `Meter matrix imported — ${rows} day-row(s) across ${assetsMatched} asset(s) over ${dates} date(s).`;
    if (resets) msg += ` ${resets} meter reset(s) handled.`;
    if (unknown.length) msg += ` Skipped unknown codes: ${unknown.join(", ")}.`;
    setStatus(msg);
    if (unknown.length) {
      alert(`Import complete, but these header codes were not found and were skipped:\n\n${unknown.join(", ")}`);
    }
    await loadDailyInput().catch(() => {});
  } catch (e) {
    setStatus("Meter matrix upload failed: " + (e.message || e));
    alert("Meter matrix upload failed: " + (e.message || e));
  }
}

/* =========================
   REPORTS
========================= */

/** Start-up: CSV upload and FAMS fuel import controls. Called once from init() in init.js. */
function wireUploadControls() {
  qs("doUpload")?.addEventListener("click", () =>
    doUpload().catch((e) => setStatus("Upload error: " + e.message))
  );
  qs("fuelFamsUploadBtn")?.addEventListener("click", () =>
    importFamsFuelFile().catch((e) => setStatus("FAMS import error: " + e.message))
  );
  qs("fuelFamsSyncNowBtn")?.addEventListener("click", () =>
    syncFamsFuelNow().catch((e) => setStatus("FAMS sync error: " + e.message))
  );
  qs("fuelFamsSyncSelectedBtn")?.addEventListener("click", () =>
    syncFamsFuelSelectedDates().catch((e) => setStatus("FAMS selected-date sync error: " + e.message))
  );
  qs("fuelFamsDuplicatesPreviewBtn")?.addEventListener("click", () =>
    previewFamsFuelDuplicates().catch((e) => setStatus("FAMS duplicate preview error: " + e.message))
  );
  qs("fuelFamsDuplicatesRemoveBtn")?.addEventListener("click", () =>
    removeFamsFuelDuplicates().catch((e) => setStatus("FAMS duplicate cleanup error: " + e.message))
  );
  qs("fuelFamsRefreshStatusBtn")?.addEventListener("click", () =>
    loadFamsFuelStatus().catch((e) => setStatus("FAMS status error: " + e.message))
  );
  qs("fuelRepairMeterChainBtn")?.addEventListener("click", () =>
    repairFuelMeterChain().catch((e) => setStatus("Meter chain repair error: " + e.message))
  );
  qs("fuelClearFromDateBtn")?.addEventListener("click", () =>
    clearFuelFromDate().catch((e) => setStatus("Fuel clear error: " + e.message))
  );
  qs("fuelClearPreviewBtn")?.addEventListener("click", () =>
    runFuelClearPreview().catch((e) => setStatus("Fuel clear preview error: " + e.message))
  );
  qs("downloadFuelTemplate")?.addEventListener("click", downloadFuelCsvTemplate);
  qs("downloadStoreTemplate")?.addEventListener("click", downloadStoresCsvTemplate);
  qs("downloadFuelBaselineTemplate")?.addEventListener("click", downloadFuelBaselineCsvTemplate);
}

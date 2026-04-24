(function () {
  function qs(id) {
    return document.getElementById(id);
  }

  function statusClass(value) {
    const v = String(value || "").toLowerCase();
    if (v === "down") return "status down";
    if (v === "production") return "status production";
    if (v === "standby") return "status standby";
    return "status unknown";
  }

  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    const text = await res.text();
    let data = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch {
      data = null;
    }
    if (!res.ok) {
      const msg = data?.error || data?.message || text || `Request failed (${res.status})`;
      throw new Error(msg);
    }
    return data || {};
  }

  async function loadQrProfile() {
    const url = new URL(window.location.href);
    const assetCode = String(url.searchParams.get("asset_code") || "").trim();
    if (!assetCode) {
      qs("sub").textContent = "Missing asset_code in QR URL.";
      return;
    }

    qs("sub").textContent = `Loading ${assetCode}...`;
    const data = await fetchJson(`/api/assets/${encodeURIComponent(assetCode)}/qr-profile`);
    const payload = data?.live_preview || data?.stored?.qr_payload || {};
    const service = payload?.next_service_due;
    const meter = payload?.meter;
    const fuel = payload?.fuel;
    const status = String(payload?.status || "UNKNOWN").toUpperCase();

    qs("sub").textContent = `Asset ${assetCode} loaded`;
    qs("machineCode").textContent = String(payload?.asset?.asset_code || assetCode);
    qs("machineMake").textContent = String(payload?.asset?.make || "-");
    qs("machineModel").textContent = String(payload?.asset?.model || "-");
    const statusEl = qs("machineStatus");
    statusEl.textContent = status;
    statusEl.className = statusClass(status);
    qs("meter").textContent = meter?.current_hours != null ? `${Number(meter.current_hours).toFixed(1)}h` : "-";
    qs("nextService").textContent = service
      ? `${service.service_name} @ ${service.next_due_hours}h (${service.remaining_hours}h remaining)`
      : "No active maintenance plan";
    qs("fuel30d").textContent = `${Number(fuel?.liters_last_30_days || 0).toFixed(1)} L`;
    qs("inspectionDate").textContent = String(payload?.inspections?.last_inspection_date || "No inspection date");
  }

  qs("refreshBtn")?.addEventListener("click", () => {
    loadQrProfile().catch((e) => {
      qs("sub").textContent = `Error: ${e.message || e}`;
    });
  });

  loadQrProfile().catch((e) => {
    qs("sub").textContent = `Error: ${e.message || e}`;
  });
})();

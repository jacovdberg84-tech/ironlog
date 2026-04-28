(function () {
  function qs(id) {
    return document.getElementById(id);
  }
  function setText(id, value) {
    const el = qs(id);
    if (!el) return;
    el.textContent = value;
  }

  function statusClass(value) {
    const v = String(value || "").toLowerCase();
    if (v === "down") return "status down";
    if (v === "production") return "status production";
    if (v === "standby") return "status standby";
    return "status unknown";
  }

  function inferMakeModelFromAsset(payload) {
    const asset = payload?.asset || {};
    const fromPayloadMake = String(asset.make || "").trim();
    const fromPayloadModel = String(asset.model || "").trim();
    if (fromPayloadMake || fromPayloadModel) {
      return {
        make: fromPayloadMake || "-",
        model: fromPayloadModel || "-",
      };
    }

    const name = String(asset.asset_name || "").trim();
    const code = String(asset.asset_code || "").trim();
    const tokens = name.split(/\s+/).filter(Boolean);
    let make = "";
    let model = "";

    if (tokens.length) make = tokens[0].toUpperCase();
    if (tokens.length >= 2) {
      const second = String(tokens[1] || "");
      if (/[0-9]/.test(second) || second.length <= 12) model = second.toUpperCase();
    }
    if (!model && code) {
      const codeToken = code.split(/[-_\s]/).find((t) => /[0-9]/.test(t));
      if (codeToken) model = codeToken.toUpperCase();
    }
    return {
      make: make || "-",
      model: model || "-",
    };
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
      setText("sub", "Missing asset_code in QR URL.");
      return;
    }

    setText("sub", `Loading ${assetCode}...`);
    const data = await fetchJson(`/api/assets/${encodeURIComponent(assetCode)}/qr-profile`);
    const payload = data?.live_preview || data?.stored?.qr_payload || {};
    const service = payload?.next_service_due;
    const meter = payload?.meter;
    const fuel = payload?.fuel;
    const status = String(payload?.status || "UNKNOWN").toUpperCase();

    setText("sub", `Asset ${assetCode} loaded`);
    setText("machineCode", String(payload?.asset?.asset_code || assetCode));
    const mm = inferMakeModelFromAsset(payload);
    setText("machineMake", mm.make);
    setText("machineModel", mm.model);
    const statusEl = qs("machineStatus");
    if (statusEl) {
      statusEl.textContent = status;
      statusEl.className = statusClass(status);
    }
    setText("meter", meter?.current_hours != null ? `${Number(meter.current_hours).toFixed(1)}h` : "-");
    setText("nextService", service
      ? `${service.service_name} @ ${service.next_due_hours}h (${service.remaining_hours}h remaining)`
      : "No active maintenance plan");
    setText("fuel30d", `${Number(fuel?.liters_last_30_days || 0).toFixed(1)} L`);
    setText("inspectionDate", String(payload?.inspections?.last_inspection_date || "No inspection date"));

    const openInspectionsBtn = qs("openInspectionsBtn");
    if (openInspectionsBtn) openInspectionsBtn.href = `./asset-qr-detail.html?view=inspections&asset_code=${encodeURIComponent(assetCode)}`;
    const openHistoryBtn = qs("openHistoryBtn");
    if (openHistoryBtn) openHistoryBtn.href = `./asset-qr-detail.html?view=history&asset_code=${encodeURIComponent(assetCode)}`;
    const openWoBtn = qs("openWoBtn");
    if (openWoBtn) openWoBtn.href = `./asset-qr-detail.html?view=workorders&asset_code=${encodeURIComponent(assetCode)}`;
  }

  qs("refreshBtn")?.addEventListener("click", () => {
    loadQrProfile().catch((e) => {
      setText("sub", `Error: ${e.message || e}`);
    });
  });

  loadQrProfile().catch((e) => {
    setText("sub", `Error: ${e.message || e}`);
  });
})();

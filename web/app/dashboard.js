// IRONLOG/web/app/dashboard.js — Dashboard KPIs.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function loadDashboard() {
  const dateEl = qs("date");
  const scheduledEl = qs("scheduled");
  const date = dateEl ? dateEl.value : new Date().toISOString().slice(0, 10);
  const scheduled = scheduledEl ? scheduledEl.value || 10 : 10;

  setStatus("Loading dashboard...");
  setSkeleton("downtimeList", 2);
  setSkeleton("downtimeReasonsList", 2);
  setSkeleton("ldvPrestartList", 2);
  setSkeleton("stockList", 2);
  setSkeleton("woList", 2);
  setSkeleton("riskBoardList", 2);
  if (isDashSectionVisible("dashSlaCard")) setSkeleton("slaList", 2);
  if (isDashSectionVisible("dashSyncDiagCard")) setSkeleton("syncDiagList", 2);
  if (isDashSectionVisible("dashCostTrendCard")) setSkeleton("costTrendList", 2);
  setSkeleton("costList", 2);
  setSkeleton("lubeList", 2);
  if (isDashSectionVisible("dashStockMonitorCard")) setSkeleton("stockMonitorList", 2);
  setSkeleton("telematicsFleetList", 2);
  setSkeleton("cartrackFleetHost", 2);

  const data = await fetchJson(`${API}/api/dashboard?date=${date}&scheduled=${scheduled}`);
  loadTelematicsFleet().catch(() => {});
  loadCartrackFleet().catch(() => {});

  const sqDateEl = qs("sqDate");
  if (sqDateEl && !sqDateEl.value) sqDateEl.value = date;

  const _kpiTh = getThresholds();
  setSpeedo(qs("availNeedle"), qs("gAvailVal"), data?.kpi?.availability, { goodAt: _kpiTh.availTarget, warnAt: _kpiTh.availCrit });
  setSpeedo(qs("utilNeedle"), qs("gUtilVal"), data?.kpi?.utilization, { goodAt: _kpiTh.utilTarget, warnAt: _kpiTh.utilCrit });
  updateKpiAlertBanner(
    data?.kpi?.availability_mtd ?? data?.kpi?.availability,
    data?.kpi?.utilization_mtd ?? data?.kpi?.utilization
  );

  const mtdRange =
    data.kpi?.mtd_start && data.kpi?.mtd_end
      ? `${data.kpi.mtd_start} → ${data.kpi.mtd_end}`
      : "";
  const k = data.kpi || {};
  const siteTag = k.site_code ? ` · Site: ${k.site_code}` : "";
  const gaugeNote =
    k.gauge_basis === "mtd"
      ? "Gauges use MTD (no planned hour-meter time on the selected day). "
      : `Gauges use selected day (${data.date || ""}). `;
  setText(
    "availMeta",
    `${gaugeNote}${siteTag}` +
      (mtdRange
        ? `MTD ${mtdRange} · Distinct assets (MTD): ${k.used_assets ?? "—"} | Planned−down (MTD) hrs: ${k.available_hours ?? "—"} | Downtime (MTD): ${k.downtime_hours ?? "—"}`
        : `Distinct assets (MTD): ${k.used_assets ?? "—"} | Planned−down (MTD) hrs: ${k.available_hours ?? "—"} | Downtime (MTD): ${k.downtime_hours ?? "—"}`)
  );
  setText(
    "utilMeta",
    `Day planned hrs: ${Number(k.scheduled_hours_day ?? 0).toFixed(1)} · Day run hrs: ${Number(k.run_hours ?? 0).toFixed(1)}. ` +
      (mtdRange
        ? `MTD ${mtdRange} · Run (MTD): ${Number(k.run_hours_mtd ?? k.run_hours ?? 0).toFixed(1)} | Planned (MTD): ${Number(k.utilization_base_hours || 0).toFixed(1)} | MTD util: ${k.utilization_mtd != null ? `${Number(k.utilization_mtd).toFixed(2)}%` : "—"} | Scheduled/asset (header): ${data.scheduled_hours_per_asset}`
        : `Run (MTD): ${Number(k.run_hours_mtd ?? k.run_hours ?? 0).toFixed(1)} | Planned (MTD): ${Number(k.utilization_base_hours || 0).toFixed(1)} | Scheduled/asset: ${data.scheduled_hours_per_asset}`)
  );
  const debugToggle = qs("kpiDebugToggle");
  const debugList = qs("kpiDebugList");
  if (debugList) {
    const show = Boolean(debugToggle?.checked);
    debugList.style.display = show ? "" : "none";
    debugList.innerHTML = "";
    if (show) {
      const rows = Array.isArray(data.per_asset_kpi) ? data.per_asset_kpi : [];
      rows.forEach((r) => {
        const mode = String(r.utilization_mode || "hours").toLowerCase();
        const meterTxt = mode === "km"
          ? `Meter: ${Number(r.meter_run_value || 0).toFixed(2)} km`
          : `Meter: ${Number(r.meter_run_value || r.run_hours || 0).toFixed(2)} h`;
        debugList.appendChild(
          item(
            `<b>${r.asset_code || `ID ${r.asset_id}`}</b>` +
            `<br><small>Mode: ${mode.toUpperCase()}${mode === "km" ? ` (km/h factor ${Number(r.km_per_hour_factor || 10).toFixed(2)})` : ""} | ${meterTxt}</small>` +
            `<br><small>Sched: ${Number(r.scheduled_hours || 0).toFixed(2)} | Down: ${Number(r.downtime_hours || 0).toFixed(2)} | Avail: ${Number(r.available_hours || 0).toFixed(2)} | Run(H): ${Number(r.run_hours || 0).toFixed(2)}${r.contributes_to_kpi === false ? " | KPI: EXCLUDED" : ""}</small>`
          )
        );
      });
      if (!rows.length) {
        debugList.appendChild(
          item(
            "<small>No hour-meter production assets for this site/date (per-asset rows = selected day; gauges = selected day; MTD figures in subtitles).</small>"
          )
        );
      }
    }
  }

  setText("aLowStock", data.alerts.low_stock);
  setText("aOverdue", data.alerts.overdue_maintenance);
  setText("aOpenWO", data.alerts.open_work_orders);

  const downtimeList = qs("downtimeList");
  if (downtimeList) {
    downtimeList.innerHTML = "";
    (data.major_downtime || []).forEach((r) => {
      downtimeList.appendChild(
        item(
          `<b>${r.asset_code}</b> – ${r.downtime_hours}h ${
            r.critical ? " <span class='pill red'>CRIT</span>" : ""
          }<br><small>${r.description}</small>`
        )
      );
    });
    if (!data.major_downtime?.length) downtimeList.appendChild(item("<small>No downtime recorded for this date.</small>"));
  }

  const reasonsList = qs("downtimeReasonsList");
  if (reasonsList) {
    reasonsList.innerHTML = "";
    (data.downtime_reasons || []).forEach((r) => {
      reasonsList.appendChild(
        item(`<b>${r.reason}</b> – ${r.hours_down}h<br><small>Incidents: ${r.incidents}</small>`)
      );
    });
    if (!data.downtime_reasons?.length) {
      reasonsList.appendChild(item("<small>No downtime reasons logged for this date.</small>"));
    }
  }

  const ldvPrestartList = qs("ldvPrestartList");
  if (ldvPrestartList) {
    ldvPrestartList.innerHTML = "";
    const badge = qs("ldvPrestartBadge");
    const card = qs("ldvPrestartCard");
    const compliantPill = qs("ldvPrestartCompliantPill");
    const missingPill = qs("ldvPrestartMissingPill");
    const pctPill = qs("ldvPrestartPctPill");
    const setComplianceTone = (pct) => {
      const p = Number(pct || 0);
      const th = getLdvPrestartThresholds();
      const tone = p >= th.greenAt ? "green" : p >= th.warnAt ? "orange" : "red";
      if (badge) {
        badge.className = `dash-card-badge${tone === "red" ? " dash-card-badge-red" : ""}`;
        badge.textContent = tone === "green" ? "On Track" : tone === "orange" ? "Attention" : "Critical";
      }
      if (card) {
        card.style.borderColor = tone === "green"
          ? "rgba(34,197,94,0.55)"
          : tone === "orange"
            ? "rgba(245,158,11,0.55)"
            : "rgba(239,68,68,0.55)";
      }
      if (compliantPill) compliantPill.className = `kpi-pill ${tone === "green" ? "kpi-pill-green" : "kpi-pill-blue"}`;
      if (missingPill) missingPill.className = `kpi-pill ${tone === "red" ? "kpi-pill-red" : "kpi-pill-orange"}`;
      if (pctPill) pctPill.className = `kpi-pill ${tone === "green" ? "kpi-pill-green" : tone === "orange" ? "kpi-pill-orange" : "kpi-pill-red"}`;
    };
    try {
      const ps = await fetchJson(`${API}/api/dashboard/ldv-prestart/compliance?date=${encodeURIComponent(date)}`);
      const summary = ps?.summary || {};
      setText("ldvPrestartCompliant", Number(summary.compliant || 0));
      setText("ldvPrestartMissing", Number(summary.missing || 0));
      setText("ldvPrestartPct", `${Number(summary.pct || 0).toFixed(1)}%`);
      setComplianceTone(Number(summary.pct || 0));
      const rows = Array.isArray(ps?.rows) ? ps.rows : [];
      const attentionRows = rows.filter((r) => r.status !== "compliant");
      const renderRows = attentionRows.length ? attentionRows : rows.slice(0, 5);
      renderRows.forEach((r) => {
        const statusPill = r.status === "compliant"
          ? "<span class='pill green'>OK</span>"
          : "<span class='pill orange'>PENDING</span>";
        ldvPrestartList.appendChild(
          item(
            `<b>${escapeHtml(String(r.asset_code || "-"))}</b> ${statusPill}` +
            `<br><small>${escapeHtml(String(r.reason || ""))}</small>`
          )
        );
      });
      if (!rows.length) {
        ldvPrestartList.appendChild(item("<small>No LDV assets found for V01AM-V15AM.</small>"));
      }
    } catch (e) {
      setText("ldvPrestartCompliant", "0");
      setText("ldvPrestartMissing", "0");
      setText("ldvPrestartPct", "-");
      setComplianceTone(0);
      if (badge) {
        badge.className = "dash-card-badge dash-card-badge-red";
        badge.textContent = "Unavailable";
      }
      ldvPrestartList.appendChild(
        item(`<small>Pre-start compliance unavailable: ${escapeHtml(String(e.message || e))}</small>`)
      );
    }
  }

  const stockList = qs("stockList");
  if (stockList) {
    stockList.innerHTML = "";
    (data.critical_low_stock || []).forEach((r) => {
      stockList.appendChild(
        item(`<b>${r.part_code}</b> – ${r.on_hand} on hand<br><small>${r.part_name} | Min: ${r.min_stock}</small>`)
      );
    });
    if (!data.critical_low_stock?.length) stockList.appendChild(item("<small>No critical low stock.</small>"));
  }

  const woList = qs("woList");
  if (woList) {
    woList.innerHTML = "";
    const isStrictOpenWO = (r) => {
      const norm = String(r?.status || "").trim().toLowerCase().replace(/\s+/g, "_");
      const completedAt = String(r?.completed_at || "").trim();
      const closedAt = String(r?.closed_at || "").trim();
      return ["open", "assigned", "in_progress"].includes(norm) && !completedAt && !closedAt;
    };
    const openRows = (data.open_work_orders || []).filter(isStrictOpenWO);
    setText("aOpenWO", openRows.length);
    openRows.forEach((r) => {
      woList.appendChild(
        item(`<b>WO #${r.id}</b> – ${r.asset_code}<br><small>${r.source} | ${r.status} | ${r.opened_at}</small>`)
      );
    });
    if (!openRows.length) woList.appendChild(item("<small>No open work orders.</small>"));
  }

  const riskBoardList = qs("riskBoardList");
  if (riskBoardList) {
    riskBoardList.innerHTML = "";
    try {
      const rb = await fetchJson(`${API}/api/ironmind/risk-board?date=${encodeURIComponent(date)}&limit=8`);
      const rows = Array.isArray(rb?.items) ? rb.items : [];
      rows.forEach((r) => {
        const reasons = Array.isArray(r.reasons) ? r.reasons.slice(0, 2).join(" | ") : "";
        riskBoardList.appendChild(
          item(
            `<b>${escapeHtml(r.asset_code || "-")}</b> - Risk ${Number(r.risk_score || 0).toFixed(0)}/100` +
            ` <span class="pill orange">Conf ${Number(r.confidence || 0).toFixed(0)}%</span>` +
            (reasons ? `<br><small>${escapeHtml(reasons)}</small>` : "") +
            `<br><button data-ironmind-risk-asset="${escapeHtml(r.asset_code || "")}">Open Asset History</button> ` +
            `<button data-ironmind-risk-wo="${escapeHtml(r.asset_code || "")}">Create WO</button>`
          )
        );
      });
      if (!rows.length) riskBoardList.appendChild(item("<small>No risk-board data yet. Refresh Borris insight first.</small>"));
    } catch (e) {
      riskBoardList.appendChild(item(`<small>Risk board unavailable: ${escapeHtml(e.message || String(e))}</small>`));
    }
  }

  const slaList = qs("slaList");
  if (isDashSectionVisible("dashSlaCard") && slaList) {
  const sla = data.workorder_sla || {};
  const slaSummary = sla.summary || {};
  setText("slaOpen24", Number(slaSummary.open_gt_24h || 0));
  setText("slaProgress48", Number(slaSummary.in_progress_gt_48h || 0));
  setText("slaCompleted12", Number(slaSummary.completed_gt_12h || 0));
  if (slaList) {
    slaList.innerHTML = "";
    (sla.breaches || []).forEach((r) => {
      const p = String(r.priority || "P3").toUpperCase();
      const pClass = p === "P1" ? "pri-p1" : p === "P2" ? "pri-p2" : "pri-p3";
      const s = String(r.status || "").toLowerCase();
      const actionBtn =
        s === "open"
          ? `<button data-sla-set-id="${r.id}" data-sla-set-status="assigned">Assign Now</button>`
          : s === "assigned"
          ? `<button data-sla-set-id="${r.id}" data-sla-set-status="in_progress">Start Now</button>`
          : s === "completed"
          ? `<button data-sla-set-id="${r.id}" data-sla-set-status="approved">Approve Now</button>`
          : "";
      slaList.appendChild(
        item(
          `<b>WO #${r.id}</b> - ${r.asset_code} (${r.status})` +
          ` <span class="pill ${pClass}">${p}</span>` +
          `<br><small>Age: ${Number(r.age_hours || 0)}h | Source: ${r.source || "-"} | Opened: ${r.opened_at || "-"}</small>` +
          `<br>${actionBtn} <button data-sla-nudge-id="${r.id}">Nudge Supervisor</button> <button data-sla-open-id="${r.id}">Open WO</button>`
        )
      );
    });
    if (!sla.breaches?.length) slaList.appendChild(item("<small>No SLA breaches right now.</small>"));
  }
  }

  if (isDashSectionVisible("dashSlaCard") && slaList && !slaList.dataset.bound) {
    slaList.dataset.bound = "1";
    slaList.addEventListener("click", async (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const setId = target.getAttribute("data-sla-set-id");
      const setStatus = target.getAttribute("data-sla-set-status");
      const nudgeId = target.getAttribute("data-sla-nudge-id");
      const openId = target.getAttribute("data-sla-open-id");

      try {
        if (setId && setStatus) {
          setStatus(`Updating WO #${setId} -> ${setStatus}...`);
          await transitionWorkOrderStatus(setId, setStatus);
          await loadDashboard();
          setStatus(`WO #${setId} moved to ${setStatus}.`);
          return;
        }
        if (nudgeId) {
          setStatus(`Sending supervisor nudge for WO #${nudgeId}...`);
          await nudgeSupervisor(nudgeId);
          setStatus(`Nudge sent for WO #${nudgeId}.`);
          return;
        }
        if (openId) {
          const url = `/web/workorders.html?wo=${encodeURIComponent(openId)}`;
          if (getSlaOpenSameTab()) {
            window.location.href = url;
          } else {
            window.open(url, "_blank");
          }
        }
      } catch (e) {
        setStatus(`SLA action failed: ${e.message || e}`);
      }
    });
  }

  const syncDiagList = qs("syncDiagList");
  const syncDiagTrend = qs("syncDiagTrend");
  if (isDashSectionVisible("dashSyncDiagCard") && syncDiagList) {
    syncDiagList.innerHTML = "";
    try {
      const sd = await fetchJson(`${API}/api/sync/diagnostics`);
      const tableRows = Array.isArray(sd?.outbox_unsynced_by_table) ? sd.outbox_unsynced_by_table : [];
      const errRows = Array.isArray(sd?.outbox_error_breakdown) ? sd.outbox_error_breakdown : [];
      const peers = Array.isArray(sd?.checkpoints) ? sd.checkpoints : [];
      const unsyncedTotal = tableRows.reduce((sum, r) => sum + Number(r?.count || 0), 0);
      const outboxErrors = errRows
        .filter((r) => String(r?.error_text || "").trim() !== "" && String(r?.error_text || "") !== "(none)")
        .reduce((sum, r) => sum + Number(r?.count || 0), 0);
      const peerCount = new Set(peers.map((r) => String(r?.peer_name || "").trim()).filter(Boolean)).size;
      setText("sdOutboxUnsynced", String(unsyncedTotal));
      setText("sdOutboxErrors", String(outboxErrors));
      setText("sdPeers", String(peerCount));

      // Keep a short local trend history to show direction.
      const trendKey = "ironlog_sync_diag_trend_v1";
      let trendRows = [];
      try {
        const raw = localStorage.getItem(trendKey);
        const parsed = raw ? JSON.parse(raw) : [];
        trendRows = Array.isArray(parsed) ? parsed : [];
      } catch {
        trendRows = [];
      }
      trendRows.push({
        t: new Date().toISOString(),
        unsynced: Number(unsyncedTotal || 0),
        errors: Number(outboxErrors || 0),
      });
      trendRows = trendRows.slice(-12);
      try {
        localStorage.setItem(trendKey, JSON.stringify(trendRows));
      } catch {
        // ignore localStorage write failures
      }

      tableRows.slice(0, 5).forEach((r) => {
        syncDiagList.appendChild(
          item(`<b>${escapeHtml(r.table_name || "-")}</b> — ${Number(r.count || 0)} pending`)
        );
      });
      errRows
        .filter((r) => String(r?.error_text || "").trim() !== "" && String(r?.error_text || "") !== "(none)")
        .slice(0, 3)
        .forEach((r) => {
          syncDiagList.appendChild(
            item(`<small>Error: ${escapeHtml(String(r.error_text || "-"))} (${Number(r.count || 0)})</small>`)
          );
        });
      if (!tableRows.length) {
        syncDiagList.appendChild(item("<small>No outbox backlog detected.</small>"));
      }

      if (syncDiagTrend) {
        const maxUnsynced = Math.max(1, ...trendRows.map((r) => Number(r.unsynced || 0)));
        const maxErrors = Math.max(1, ...trendRows.map((r) => Number(r.errors || 0)));
        const pointsUnsynced = trendRows
          .map((r, i) => {
            const x = trendRows.length <= 1 ? 0 : (i / (trendRows.length - 1)) * 100;
            const y = 100 - (Number(r.unsynced || 0) / maxUnsynced) * 100;
            return `${x.toFixed(2)},${y.toFixed(2)}`;
          })
          .join(" ");
        const pointsErrors = trendRows
          .map((r, i) => {
            const x = trendRows.length <= 1 ? 0 : (i / (trendRows.length - 1)) * 100;
            const y = 100 - (Number(r.errors || 0) / maxErrors) * 100;
            return `${x.toFixed(2)},${y.toFixed(2)}`;
          })
          .join(" ");
        syncDiagTrend.innerHTML = `
          <div class="row" style="justify-content:space-between; align-items:center; margin-bottom:6px;">
            <small class="muted">Backlog trend (last ${trendRows.length} samples)</small>
            <small class="muted">Unsynced <span style="color:#2563eb;">●</span> Errors <span style="color:#dc2626;">●</span></small>
          </div>
          <svg viewBox="0 0 100 100" preserveAspectRatio="none" style="width:100%; height:90px; background:var(--bg-f8fafc); border:1px solid var(--bd-e5e7eb); border-radius:8px;">
            <polyline fill="none" stroke="#2563eb" stroke-width="2.2" points="${pointsUnsynced}"></polyline>
            <polyline fill="none" stroke="#dc2626" stroke-width="2.2" points="${pointsErrors}"></polyline>
          </svg>
        `;
      }
    } catch (e) {
      setText("sdOutboxUnsynced", "0");
      setText("sdOutboxErrors", "0");
      setText("sdPeers", "0");
      const msg = String(e?.message || "");
      if (msg.toLowerCase().includes("403")) {
        syncDiagList.appendChild(item("<small>Sync diagnostics available to admin/supervisor roles.</small>"));
      } else {
        syncDiagList.appendChild(item(`<small>Sync diagnostics unavailable: ${escapeHtml(msg || "unknown error")}</small>`));
      }
      if (syncDiagTrend) syncDiagTrend.innerHTML = "";
    }
  }

  const refreshSyncBtn = qs("refreshSyncDiagnostics");
  if (isDashSectionVisible("dashSyncDiagCard") && refreshSyncBtn && !refreshSyncBtn.dataset.bound) {
    refreshSyncBtn.dataset.bound = "1";
    refreshSyncBtn.addEventListener("click", () => {
      loadDashboard().catch((e) => setStatus(`Dashboard reload failed: ${e.message || e}`));
    });
  }

  const relDays = Number(qs("relDays")?.value || 30);
  const relRange = getLastNDaysRange(date, relDays);
  const rel = await fetchJson(
    `${API}/api/dashboard/reliability?start=${encodeURIComponent(relRange.start)}&end=${encodeURIComponent(relRange.end)}`
  );
  setText("relMtbf", rel.mtbf_hours == null ? "-" : Number(rel.mtbf_hours).toFixed(2));
  setText("relLttr", rel.lttr_hours == null ? "-" : Number(rel.lttr_hours).toFixed(2));
  setText("relFailures", String(Number(rel.failure_count || 0)));
  setText("relWindow", `Window: ${relRange.start} to ${relRange.end} (${Math.max(1, relDays)} days)`);
  const relTrend = await fetchJson(
    `${API}/api/dashboard/reliability/trend?weeks=12&end=${encodeURIComponent(date)}`
  );
  const relPoints = Array.isArray(relTrend?.points) ? relTrend.points : [];
  const relChart = qs("relTrendChart");
  const relList = qs("relTrendList");
  if (relChart) {
    relChart.innerHTML = "";
    const maxMtbf = Math.max(1, ...relPoints.map((p) => Number(p.mtbf_hours || 0)));
    relPoints.forEach((p) => {
      const bar = document.createElement("div");
      bar.className = "cost-bar";
      const h = Math.max(6, Math.round((Number(p.mtbf_hours || 0) / maxMtbf) * 100));
      bar.style.height = `${h}px`;
      bar.title = `${p.start} to ${p.end} | MTBF ${p.mtbf_hours ?? "-"} | LTTR ${p.lttr_hours ?? "-"} | Failures ${p.failure_count || 0}`;
      bar.innerHTML =
        `<span class="cost-bar-value">${p.mtbf_hours == null ? "-" : Number(p.mtbf_hours).toFixed(1)}</span>` +
        `<span class="cost-bar-label">${p.label || ""}</span>`;
      relChart.appendChild(bar);
    });
    if (!relPoints.length) relChart.appendChild(item("<small>No reliability trend data.</small>"));
  }
  if (relList) {
    relList.innerHTML = "";
    relPoints.slice(-6).reverse().forEach((p) => {
      relList.appendChild(
        item(
          `<b>${p.start} to ${p.end}</b>` +
          `<br><small>MTBF: ${p.mtbf_hours == null ? "-" : Number(p.mtbf_hours).toFixed(2)} | LTTR: ${p.lttr_hours == null ? "-" : Number(p.lttr_hours).toFixed(2)} | Failures: ${Number(p.failure_count || 0)}</small>`
        )
      );
    });
  }

  let trendRows = [];
  if (isDashSectionVisible("dashCostTrendCard")) {
  const trend = await fetchJson(`${API}/api/dashboard/cost/trend?months=12`);
  trendRows = Array.isArray(trend.rows) ? trend.rows : [];
  const mom = trend.mom || {};
  setText("ctCurrentMonth", trend.latest?.month || "-");
  setText("ctCurrentTotal", fmtMoney(trend.latest?.total_cost || 0));
  if (mom.variance == null) {
    setText("ctMoM", "N/A");
  } else {
    const pct = mom.variance_pct == null ? "" : ` (${Number(mom.variance_pct).toFixed(1)}%)`;
    setText("ctMoM", `${Number(mom.variance) >= 0 ? "+" : ""}${fmtMoney(mom.variance)}${pct}`);
  }
  const trendList = qs("costTrendList");
  const trendChart = qs("costTrendChart");
  if (trendChart) {
    trendChart.innerHTML = "";
    const maxCost = trendRows.reduce((m, r) => Math.max(m, Number(r.total_cost || 0)), 0);
    trendRows.forEach((r, idx) => {
      const total = Number(r.total_cost || 0);
      const h = maxCost > 0 ? Math.max(8, Math.round((total / maxCost) * 92)) : 8;
      const bar = document.createElement("div");
      const prev = idx > 0 ? Number(trendRows[idx - 1]?.total_cost || 0) : null;
      let trendClass = "neutral";
      if (prev != null && Number.isFinite(prev)) {
        if (total > prev) trendClass = "up";
        else if (total < prev) trendClass = "down";
      }
      bar.className = `cost-bar ${trendClass}`;
      bar.style.height = `${h}px`;
      bar.title = `${r.month}: ${fmtMoney(total)}`;

      const monthLabel = document.createElement("span");
      monthLabel.className = "cost-bar-label";
      monthLabel.textContent = String(r.month || "").slice(5);

      const valueLabel = document.createElement("span");
      valueLabel.className = "cost-bar-value";
      valueLabel.textContent = fmtMoney(total);

      bar.appendChild(monthLabel);
      bar.appendChild(valueLabel);
      trendChart.appendChild(bar);
    });
    if (!trendRows.length) trendChart.innerHTML = "<small class='muted'>No trend data.</small>";
  }
  if (trendList) {
    trendList.innerHTML = "";
    trendRows.slice().reverse().forEach((r) => {
      trendList.appendChild(
        item(
          `<b>${r.month}</b> - ${fmtMoney(r.total_cost)}` +
          `<br><small>Fuel ${fmtMoney(r.fuel_cost)} | Lube ${fmtMoney(r.lube_cost)} | Parts ${fmtMoney(r.parts_cost)} | Labor ${fmtMoney(r.labor_cost)} | Down ${fmtMoney(r.downtime_cost)}</small>`
        )
      );
    });
    if (!trendRows.length) trendList.appendChild(item("<small>No monthly cost trend data.</small>"));
  }
  }

  const costs = data.cost_engine || {};
  setText("cTotalCost", fmtMoney(costs.total_cost));
  setText("cCostPerHour", costs.cost_per_run_hour == null ? "N/A" : fmtMoney(costs.cost_per_run_hour));
  setText("cLaborHours", Number(costs.labor_hours || 0).toFixed(1));
  setText("cFuelCost", fmtMoney(costs.fuel_cost));
  setText("cLubeCost", fmtMoney(costs.lube_cost));
  setText("cPartsCost", fmtMoney(costs.parts_cost));
  setText("cLaborCost", fmtMoney(costs.labor_cost));
  setText("cDowntimeCost", fmtMoney(costs.downtime_cost));

  const costList = qs("costList");
  if (costList) {
    costList.innerHTML = "";
    (costs.top_asset_costs || []).forEach((r) => {
      costList.appendChild(
        item(
          `<b>${r.asset_code}</b> - ${fmtMoney(r.total_cost)}` +
          `<br><small>${r.asset_name || ""} | Fuel ${fmtMoney(r.fuel_cost)} | Lube ${fmtMoney(r.lube_cost)} | Parts ${fmtMoney(r.parts_cost)} | Labor ${fmtMoney(r.labor_cost)} | Down ${fmtMoney(r.downtime_cost)}</small>`
        )
      );
    });
    if (!costs.top_asset_costs?.length) costList.appendChild(item("<small>No cost activity for this date.</small>"));
  }

  const lube = data.lube_usage || {};
  setText("lubeQtyTotal", Number(lube.qty_total || 0).toFixed(1));
  setText("lubeEntries", Array.isArray(lube.rows) ? lube.rows.length : 0);
  setText("lubeAssets", Array.isArray(lube.rows) ? lube.rows.length : 0);
  setText("cLubeCost", fmtMoney(lube.total_lube_cost != null ? lube.total_lube_cost : data?.cost_engine?.lube_cost));
  lubeUsageCache = {
    rows: Array.isArray(lube.rows) ? lube.rows.map((r) => ({
      ...r,
      qty_total: Number(r.qty ?? 0),
      total_lube_cost: Number(r.total_lube_cost ?? r.lube_cost ?? 0),
    })) : [],
  };
  renderLubeUsageTable(lubeUsageCache);

  if (isDashSectionVisible("dashStockMonitorCard")) {
    await loadStockMonitor().catch(() => {});
  }
  await loadIronmindInsight({ silent: true }).catch(() => {});
  await loadIronmindHealth().catch(() => {});
  await loadIronmindSettings().catch(() => {});
  await loadIronmindHistory({ silent: true }).catch(() => {});

  setStatus("Dashboard ready.");
}

/** Start-up: Dashboard refresh and KPI controls. Called once from init() in init.js. */
function wireDashboardControls() {
  qs("refresh")?.addEventListener("click", () =>
    loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message))
  );
  qs("kpiDebugToggle")?.addEventListener("change", () =>
    loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message))
  );
  qs("loadReliability")?.addEventListener("click", () =>
    loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message))
  );
}

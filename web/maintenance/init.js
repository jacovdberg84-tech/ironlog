// IRONLOG/web/maintenance/init.js — Dark mode and page start-up wiring.
// Part of maintenance.html; the page loads these files in order and they share one global scope.

function initMaintDarkMode() {
  const toggle = document.getElementById("darkModeToggle");
  if (!toggle) return;
  
  const saved = localStorage.getItem("ironlog-theme");
  if (saved === "dark") {
    document.documentElement.setAttribute("data-theme", "dark");
  }
  
  toggle.addEventListener("click", () => {
    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    if (isDark) {
      document.documentElement.removeAttribute("data-theme");
      localStorage.setItem("ironlog-theme", "light");
    } else {
      document.documentElement.setAttribute("data-theme", "dark");
      localStorage.setItem("ironlog-theme", "dark");
    }
  });
}

document.addEventListener("DOMContentLoaded", () => {
  console.log("Maintenance UI loaded");

  renderSessionRoleLabel();

  // Dark mode toggle
  initMaintDarkMode();

  // Sidebar navigation
  initMaintSidebar();
  // Fallback delegated handler: ensures section switching always works
  // even if direct listeners are interrupted by future sidebar markup edits.
  document.addEventListener("click", (evt) => {
    const item = evt.target instanceof HTMLElement ? evt.target.closest(".nav-item[data-section]") : null;
    if (!item) return;
    const section = String(item.getAttribute("data-section") || "").trim();
    if (!isMaintenanceSection(section)) return;
    evt.preventDefault();
    scrollToSection(section);
  });
  
  // Initialize collapsible cards
  initMaintCollapsibles();

  const generateBtn = document.getElementById("generateBtn");
  const savePlanBtn = document.getElementById("savePlanBtn");
  const saveBackfillBtn = document.getElementById("saveBackfillBtn");
  const inspectproRefreshStatusBtn = document.getElementById("inspectproRefreshStatusBtn");
  const backfillBody = document.getElementById("backfillBody");
  const plansList = document.getElementById("plansList");
  const useLiveEl = document.getElementById("planUseLiveForLastService");

  syncDueThresholdInput();
  if (generateBtn) {
    generateBtn.addEventListener("click", generateWO);
  }
  document.getElementById("applyDueThresholdBtn")?.addEventListener("click", async () => {
    const v = getDueThresholdHours();
    localStorage.setItem(MAINT_DUE_THRESHOLD_KEY, String(v));
    await loadPlans();
    await loadDue();
  });
  document.getElementById("akpPlansSelectAllBtn")?.addEventListener("click", selectAllDuePlansInTable);
  document.getElementById("akpPlansClearSelBtn")?.addEventListener("click", clearDuePlanSelection);
  document.getElementById("maintenancePlanSearch")?.addEventListener("input", renderMaintenancePlanQueue);
  document.querySelectorAll("[data-maint-plan-filter]").forEach((button) => {
    button.addEventListener("click", () => setMaintenancePlanFilter(button.getAttribute("data-maint-plan-filter")));
  });
  document.getElementById("refreshMaintenancePlanningBtn")?.addEventListener("click", async () => {
    await Promise.all([loadPlans(), loadDue(), loadHistory()]);
  });
  document.getElementById("shareUpcomingServicesPdfBtn")?.addEventListener("click", () => openUpcomingServicesPdf(false, true));
  document.getElementById("downloadMaintenancePlansXlsxBtn")?.addEventListener("click", () => downloadMaintenancePlansXlsx());
  document.getElementById("openUpcomingServicesPdfBtn")?.addEventListener("click", () => openUpcomingServicesPdf(false));
  document.getElementById("downloadUpcomingServicesPdfBtn")?.addEventListener("click", () => openUpcomingServicesPdf(true));

  if (savePlanBtn) {
    savePlanBtn.addEventListener("click", savePlan);
  }
  if (saveBackfillBtn) {
    saveBackfillBtn.addEventListener("click", saveBackfillHistory);
  }
  if (inspectproRefreshStatusBtn) {
    inspectproRefreshStatusBtn.addEventListener("click", () => {
      loadInspectproStatus().catch(() => {});
    });
  }
  document.getElementById("refreshBackfillListBtn")?.addEventListener("click", () => {
    loadBackfillHistory().catch((err) => alert(err.message || err));
  });
  document.getElementById("backfillListAsset")?.addEventListener("change", () => {
    hideBackfillEditPanel();
    loadBackfillHistory().catch(() => {});
  });
  document.getElementById("backfillEditSaveBtn")?.addEventListener("click", () => {
    saveBackfillEditPanel().catch((err) => alert(err.message || err));
  });
  document.getElementById("backfillEditCancelBtn")?.addEventListener("click", hideBackfillEditPanel);
  document.getElementById("histAssetFilter")?.addEventListener("change", () => loadHistory());
  if (backfillBody) {
    backfillBody.addEventListener("click", (e) => {
      const btn = e.target instanceof HTMLElement ? e.target.closest("button[data-backfill-action]") : null;
      if (!btn) return;
      const tr = btn.closest("tr[data-backfill-id]");
      const id = Number(tr?.getAttribute("data-backfill-id") || 0);
      const action = String(btn.getAttribute("data-backfill-action") || "");
      if (!id || !action) return;
      if (action === "edit") {
        startEditBackfillHistory(id);
      } else if (action === "delete") {
        deleteBackfillHistory(id).catch((err) => alert(err.message || err));
      }
    });
  }
  if (plansList) {
    plansList.addEventListener("click", (e) => {
      const btn = e.target instanceof HTMLElement ? e.target.closest("button[data-plan-rebase-id]") : null;
      if (!btn) return;
      const row = btn.closest("tr");
      const directId = Number(btn.getAttribute("data-plan-rebase-id") || 0);
      const rowId = Number(row?.getAttribute("data-plan-id") || 0);
      const planId = directId || rowId;
      const assetId = Number(row?.getAttribute("data-plan-asset-id") || 0);
      rebasePlanLastServiceHours(planId, { asset_id: assetId }).catch((err) => alert(err.message || err));
    });
  }

  const planServiceTypeEl = document.getElementById("planServiceType");
  if (planServiceTypeEl) {
    planServiceTypeEl.addEventListener("change", syncPlanServiceTypeFromDropdown);
  }

  const assetEl = document.getElementById("planAsset");
  if (assetEl) {
    assetEl.addEventListener("change", () => {
      refreshPlanServiceTypeDropdown();
      loadLiveHoursForSelectedAsset();
    });
  }

  if (useLiveEl) {
    useLiveEl.addEventListener("change", syncLastServiceHoursFromLive);
  }

  loadAssetsForPlan().then(() => {
    refreshPlanServiceTypeDropdown();
    loadLiveHoursForSelectedAsset();
    const urlCode = String(new URLSearchParams(window.location.search).get("asset_code") || "")
      .trim()
      .toUpperCase();
    if (urlCode) openServiceRecordsForAsset(urlCode);
  });
  loadReliabilityAssets().catch(() => {});

  loadPlans();
  loadDue();
  loadHistory();
  loadBackfillHistory();
  if (!document.getElementById("inspectproStatusCard")?.classList.contains("hidden")) {
    loadInspectproStatus();
    setInterval(() => {
      loadInspectproStatus().catch(() => {});
    }, 20000);
  }
  const miDate = document.getElementById("miDate");
  const aiDate = document.getElementById("aiDate");
  const aiFormNo = document.getElementById("aiFormNumber");
  const drDate = document.getElementById("drDate");
  if (miDate && !miDate.value) miDate.value = new Date().toISOString().slice(0, 10);
  if (aiDate && !aiDate.value) aiDate.value = new Date().toISOString().slice(0, 10);
  if (aiFormNo && !aiFormNo.value) aiFormNo.value = generateArtisanFormNumber();
  if (drDate && !drDate.value) drDate.value = new Date().toISOString().slice(0, 10);
  const miStart = document.getElementById("miStart");
  const miEnd = document.getElementById("miEnd");
  const aiStart = document.getElementById("aiStart");
  const aiEnd = document.getElementById("aiEnd");
  const drStart = document.getElementById("drStart");
  const drEnd = document.getElementById("drEnd");
  if (miStart && !miStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    miStart.value = d.toISOString().slice(0, 10);
  }
  if (miEnd && !miEnd.value) miEnd.value = new Date().toISOString().slice(0, 10);
  if (aiStart && !aiStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    aiStart.value = d.toISOString().slice(0, 10);
  }
  if (aiEnd && !aiEnd.value) aiEnd.value = new Date().toISOString().slice(0, 10);
  if (drStart && !drStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    drStart.value = d.toISOString().slice(0, 10);
  }
  if (drEnd && !drEnd.value) drEnd.value = new Date().toISOString().slice(0, 10);
  const wfStart = document.getElementById("wfStart");
  const wfEnd = document.getElementById("wfEnd");
  if (wfStart && !wfStart.value) {
    const d = new Date();
    const day = d.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + mondayOffset);
    wfStart.value = d.toISOString().slice(0, 10);
  }
  if (wfEnd && !wfEnd.value) {
    const d = new Date();
    const day = d.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + mondayOffset + 6);
    wfEnd.value = d.toISOString().slice(0, 10);
  }
  const akpStart = document.getElementById("akpStart");
  const akpEnd = document.getElementById("akpEnd");
  if (akpStart && !akpStart.value && wfStart?.value) akpStart.value = wfStart.value;
  if (akpEnd && !akpEnd.value && wfEnd?.value) akpEnd.value = wfEnd.value;
  if (akpStart && !akpStart.value) {
    const d = new Date();
    const day = d.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + mondayOffset);
    akpStart.value = d.toISOString().slice(0, 10);
  }
  if (akpEnd && !akpEnd.value) {
    const d = new Date();
    const day = d.getDay();
    const mondayOffset = day === 0 ? -6 : 1 - day;
    d.setDate(d.getDate() + mondayOffset + 4);
    akpEnd.value = d.toISOString().slice(0, 10);
  }
  const relStart = document.getElementById("relStart");
  const relEnd = document.getElementById("relEnd");
  if (relStart && !relStart.value && akpStart?.value) relStart.value = akpStart.value;
  if (relEnd && !relEnd.value && akpEnd?.value) relEnd.value = akpEnd.value;
  if (relStart && !relStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 29);
    relStart.value = d.toISOString().slice(0, 10);
  }
  if (relEnd && !relEnd.value) relEnd.value = new Date().toISOString().slice(0, 10);
  const wfActionDate = document.getElementById("wfActionDate");
  if (wfActionDate && !wfActionDate.value) wfActionDate.value = new Date().toISOString().slice(0, 10);
  const histViewMode = document.getElementById("histViewMode");
  const histClosestLimitWrap = document.getElementById("histClosestLimitWrap");
  if (histClosestLimitWrap) histClosestLimitWrap.style.display = String(histViewMode?.value || "all") === "closest" ? "" : "none";
  const histEventDate = document.getElementById("histEventDate");
  if (histEventDate && !histEventDate.value) histEventDate.value = new Date().toISOString().slice(0, 10);
  const histFilterStart = document.getElementById("histFilterStart");
  const histFilterEnd = document.getElementById("histFilterEnd");
  if (histFilterStart && !histFilterStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 30);
    histFilterStart.value = d.toISOString().slice(0, 10);
  }
  if (histFilterEnd && !histFilterEnd.value) histFilterEnd.value = new Date().toISOString().slice(0, 10);
  const mpWeekStart = document.getElementById("mpWeekStart");
  const mpWeekEnd = document.getElementById("mpWeekEnd");
  const mpMonth = document.getElementById("mpMonth");
  const kpiPackMonth = document.getElementById("kpiPackMonth");
  const kpiPackSiteCodes = document.getElementById("kpiPackSiteCodes");
  const insightsStart = document.getElementById("insightsStart");
  const insightsEnd = document.getElementById("insightsEnd");
  const insightsNear = document.getElementById("insightsNearDueHours");
  const insightsHorizon = document.getElementById("insightsPredictiveHorizonHours");
  const insightsFail = document.getElementById("insightsChecklistFailThreshold");
  const insightsFuel = document.getElementById("insightsFuelVarianceThreshold");
  if (mpWeekStart && !mpWeekStart.value) {
    const w = mpWeekRangeLabel();
    mpWeekStart.value = w.start;
  }
  if (mpWeekEnd && !mpWeekEnd.value) {
    const w = mpWeekRangeLabel();
    mpWeekEnd.value = w.end;
  }
  if (mpWeekStart && wfStart?.value) mpWeekStart.value = wfStart.value;
  if (mpWeekEnd && wfEnd?.value) mpWeekEnd.value = wfEnd.value;
  if (mpMonth && !mpMonth.value) mpMonth.value = mpMonthLabel();
  if (kpiPackMonth && !kpiPackMonth.value) kpiPackMonth.value = mpMonthLabel();
  if (kpiPackSiteCodes && !kpiPackSiteCodes.value) kpiPackSiteCodes.value = "main";
  if (insightsStart && !insightsStart.value) {
    const d = new Date();
    d.setDate(d.getDate() - 29);
    insightsStart.value = d.toISOString().slice(0, 10);
  }
  if (insightsEnd && !insightsEnd.value) insightsEnd.value = new Date().toISOString().slice(0, 10);
  const savedThresholds = getInsightsThresholds();
  if (insightsNear) insightsNear.value = String(savedThresholds.near_due_hours);
  if (insightsHorizon) insightsHorizon.value = String(savedThresholds.predictive_horizon_hours);
  if (insightsFail) insightsFail.value = String(savedThresholds.checklist_fail_threshold);
  if (insightsFuel) insightsFuel.value = String(savedThresholds.fuel_variance_threshold);
  renderManagerInspectionChecklist();
  renderArtisanInspectionChecklist();
  const miPartsBody = document.getElementById("miPartsBody");
  if (miPartsBody && !miPartsBody.querySelector("tr")) addManagerInspectionPartRow();
  document.getElementById("miPullLiveHoursBtn")?.addEventListener("click", () =>
    pullManagerInspectionLiveHours().catch((e) => console.error(e))
  );
  document.getElementById("miAddPartRowBtn")?.addEventListener("click", () => addManagerInspectionPartRow());
  const miReloadHours = () => {
    const id = Number(document.getElementById("miAsset")?.value || 0);
    if (!id) {
      const meta = document.getElementById("miLiveMeta");
      const inp = document.getElementById("miMachineHours");
      if (meta) meta.textContent = "";
      if (inp) inp.value = "";
      return;
    }
    pullManagerInspectionLiveHours().catch((e) => console.error(e));
  };
  document.getElementById("miAsset")?.addEventListener("change", miReloadHours);
  document.getElementById("miDate")?.addEventListener("change", miReloadHours);
  document.getElementById("miInspectionType")?.addEventListener("change", () => {
    renderManagerInspectionChecklist();
  });
  const aiReloadHours = () => {
    const id = Number(document.getElementById("aiAsset")?.value || 0);
    if (!id) {
      const meta = document.getElementById("aiLiveMeta");
      const inp = document.getElementById("aiMachineHours");
      if (meta) meta.textContent = "";
      if (inp) inp.value = "";
      return;
    }
    pullArtisanInspectionLiveHours().catch((e) => console.error(e));
  };
  document.getElementById("aiPullLiveHoursBtn")?.addEventListener("click", () =>
    pullArtisanInspectionLiveHours().catch((e) => console.error(e))
  );
  document.getElementById("aiAsset")?.addEventListener("change", aiReloadHours);
  document.getElementById("aiDate")?.addEventListener("change", aiReloadHours);
  document.getElementById("openAiBlankPdfBtn")?.addEventListener("click", () => openArtisanBlankFormPdf(false));
  document.getElementById("downloadAiBlankPdfBtn")?.addEventListener("click", () => openArtisanBlankFormPdf(true));
  document.getElementById("miDictateBtn")?.addEventListener("click", () => appendVoiceToField("miNotes"));
  document.getElementById("aiDictateBtn")?.addEventListener("click", () => appendVoiceToField("aiNotes"));
  document.getElementById("saveMiBtn")?.addEventListener("click", saveManagerInspection);
  document.getElementById("saveAiBtn")?.addEventListener("click", saveArtisanInspection);
  document.getElementById("scanOpenAssetQrBtn")?.addEventListener("click", () => {
    const code = resolveScanAssetCode();
    if (!code) return alert("Scan or select an asset first.");
    window.open(`asset-qr.html?asset_code=${encodeURIComponent(code)}`, "_blank");
  });
  document.getElementById("scanOpenMaintenanceBtn")?.addEventListener("click", () => {
    const code = resolveScanAssetCode();
    if (!code) return alert("Scan or select an asset first.");
    const sel = document.getElementById("miFilterAsset");
    if (sel) {
      const hit = Array.from(sel.options || []).find((o) => String(o.textContent || "").toUpperCase().startsWith(`${code} -`));
      if (hit) sel.value = String(hit.value || "");
    }
    loadManagerInspections().catch(() => {});
  });
  document.getElementById("scanOpenWorkOrdersBtn")?.addEventListener("click", () => {
    const code = resolveScanAssetCode();
    if (!code) return alert("Scan or select an asset first.");
    window.open(`workorders.html?search=${encodeURIComponent(code)}`, "_blank");
  });
  document.getElementById("saveDrBtn")?.addEventListener("click", saveDamageReport);
  document.getElementById("loadMiBtn")?.addEventListener("click", loadManagerInspections);
  document.getElementById("loadAiBtn")?.addEventListener("click", loadArtisanInspections);
  document.getElementById("loadDrBtn")?.addEventListener("click", loadDamageReports);
  document.getElementById("openDrBulkPdfBtn")?.addEventListener("click", () => openDamageReportsBulkPdf(false));
  document.getElementById("downloadDrBulkPdfBtn")?.addEventListener("click", () => openDamageReportsBulkPdf(true));
  document.getElementById("downloadDrBulkXlsxBtn")?.addEventListener("click", () => downloadDamageReportsXlsx());
  document.getElementById("openDrBulkPdfPhotosBtn")?.addEventListener("click", () => openDamageReportsBulkPdf(false, true));
  document.getElementById("downloadDrBulkPdfPhotosBtn")?.addEventListener("click", () => openDamageReportsBulkPdf(true, true));
  document.getElementById("openMiBulkPdfBtn")?.addEventListener("click", () => openManagerInspectionsBulkPdf(false));
  document.getElementById("downloadMiBulkPdfBtn")?.addEventListener("click", () => openManagerInspectionsBulkPdf(true));
  document.getElementById("openMiBulkPdfPhotosBtn")?.addEventListener("click", () => openManagerInspectionsBulkPdf(false, true));
  document.getElementById("downloadMiBulkPdfPhotosBtn")?.addEventListener("click", () => openManagerInspectionsBulkPdf(true, true));
  document.getElementById("miList")?.addEventListener("click", (evt) => {
    const openPdf = evt.target?.closest?.("button[data-mi-open-pdf]");
    if (openPdf) {
      const id = Number(openPdf.getAttribute("data-mi-open-pdf") || 0);
      if (id) openManagerInspectionPdf(id, false);
      return;
    }
    const dlPdf = evt.target?.closest?.("button[data-mi-download-pdf]");
    if (dlPdf) {
      const id = Number(dlPdf.getAttribute("data-mi-download-pdf") || 0);
      if (id) openManagerInspectionPdf(id, true);
      return;
    }
    const delBtn = evt.target?.closest?.("button[data-mi-delete]");
    if (delBtn) {
      const id = Number(delBtn.getAttribute("data-mi-delete") || 0);
      if (id) deleteManagerInspection(id).catch((e) => alert(`Delete failed: ${e.message || e}`));
      return;
    }
    const btn = evt.target?.closest?.("button[data-mi-upload]");
    if (!btn) return;
    const id = Number(btn.getAttribute("data-mi-upload") || 0);
    if (!id) return;
    uploadInspectionPhoto(id).catch((e) => alert(`Photo upload failed: ${e.message || e}`));
  });
  document.getElementById("aiList")?.addEventListener("click", (evt) => {
    const openPdf = evt.target?.closest?.("button[data-ai-open-pdf]");
    if (openPdf) {
      const id = Number(openPdf.getAttribute("data-ai-open-pdf") || 0);
      if (id) openArtisanInspectionPdf(id, false);
      return;
    }
    const dlPdf = evt.target?.closest?.("button[data-ai-download-pdf]");
    if (dlPdf) {
      const id = Number(dlPdf.getAttribute("data-ai-download-pdf") || 0);
      if (id) openArtisanInspectionPdf(id, true);
    }
  });
  document.getElementById("drList")?.addEventListener("click", (evt) => {
    const openPdf = evt.target?.closest?.("button[data-dr-open-pdf]");
    if (openPdf) {
      const id = Number(openPdf.getAttribute("data-dr-open-pdf") || 0);
      if (id) openDamageReportPdf(id, false);
      return;
    }
    const dlPdf = evt.target?.closest?.("button[data-dr-download-pdf]");
    if (dlPdf) {
      const id = Number(dlPdf.getAttribute("data-dr-download-pdf") || 0);
      if (id) openDamageReportPdf(id, true);
      return;
    }
    const btn = evt.target?.closest?.("button[data-dr-upload]");
    if (!btn) return;
    const id = Number(btn.getAttribute("data-dr-upload") || 0);
    if (!id) return;
    uploadDamagePhoto(id).catch((e) => alert(`Damage photo upload failed: ${e.message || e}`));
  });
  const applyDeepLink = () => {
    try {
      const q = new URLSearchParams(window.location.search);
      const section = String(q.get("section") || "").trim();
      const assetCode = String(q.get("asset_code") || "").trim().toUpperCase();
      if (section) scrollToSection(section);
      if (assetCode) {
        const scan = document.getElementById("scanAssetCode");
        if (scan) scan.value = assetCode;
        const sel = document.getElementById("miFilterAsset");
        if (sel) {
          const hit = Array.from(sel.options || []).find((o) => String(o.textContent || "").toUpperCase().startsWith(`${assetCode} -`));
          if (hit) sel.value = String(hit.value || "");
        }
        loadManagerInspections().catch(() => {});
      }
    } catch {}
  };
  loadAssetsForInspection().then(applyDeepLink).catch(() => {});
  loadManagerInspections().catch(() => {});
  loadArtisanInspections().catch(() => {});
  loadDamageReports().catch(() => {});
  loadWeeklyForumSummary().catch(() => {});
  loadWeeklyForumReviews().catch(() => {});
  loadWeeklyForumActions().catch(() => {});
  loadWeeklyForumParts().catch(() => {});
  loadWeeklyForumInputs().catch(() => {});
  loadMaintenancePackStatus().catch(() => {});
  if (!document.getElementById("rsgProfilesCard")?.classList.contains("hidden")) {
    loadRsgProfiles().catch(() => {});
  }
  document.getElementById("syncStatsBtn")?.addEventListener("click", syncLoadStats);
  document.getElementById("syncStateBtn")?.addEventListener("click", syncLoadState);
  document.getElementById("syncOutboxBtn")?.addEventListener("click", syncLoadOutbox);
  document.getElementById("syncPullBtn")?.addEventListener("click", syncPullEvents);
  document.getElementById("syncExportLastPullBtn")?.addEventListener("click", syncExportLastPull);
  document.getElementById("syncApplyBtn")?.addEventListener("click", syncApplyLastPull);
  document.getElementById("syncApplyDryRunBtn")?.addEventListener("click", syncApplyLastPullDryRun);
  document.getElementById("syncAckBtn")?.addEventListener("click", syncAckLastPull);
  document.getElementById("syncCheckpointBtn")?.addEventListener("click", syncCheckpointLastPull);
  document.getElementById("showMainMaintBtn")?.addEventListener("click", () => scrollToSection("maintenance"));
  document.getElementById("showManagerInspectionsBtn")?.addEventListener("click", () => setTopView("mi"));
  document.getElementById("showArtisanInspectionsBtn")?.addEventListener("click", () => setTopView("ai"));
  document.getElementById("showWeeklyForumBtn")?.addEventListener("click", () => setTopView("wf"));
  document.getElementById("showTyreInspectionsBtn")?.addEventListener("click", () => setTopView("tyre"));
  document.getElementById("showAssetKpiBtn")?.addEventListener("click", () => setTopView("kpi"));
  document.getElementById("showHistogramBtn")?.addEventListener("click", () => setTopView("hist"));
  document.getElementById("showSyncAdminBtn")?.addEventListener("click", () => setTopView("sync"));
  document.getElementById("histViewMode")?.addEventListener("change", () => loadHistory());
  document.getElementById("histClosestLimit")?.addEventListener("change", () => loadHistory());
  document.getElementById("saveHistogramBtn")?.addEventListener("click", () => saveHistogramEvent());
  document.getElementById("loadHistogramBtn")?.addEventListener("click", () => loadHistogramEvents());
  document.getElementById("openHistogramPdfBtn")?.addEventListener("click", () => openHistogramPdf(false));
  document.getElementById("downloadHistogramPdfBtn")?.addEventListener("click", () => openHistogramPdf(true));
  document.getElementById("histEventBody")?.addEventListener("click", (evt) => {
    const editBtn = evt.target?.closest?.("button[data-hist-edit]");
    if (editBtn) {
      editHistogramEvent(Number(editBtn.getAttribute("data-hist-edit") || 0));
      return;
    }
    const delBtn = evt.target?.closest?.("button[data-hist-del]");
    if (delBtn) {
      deleteHistogramEvent(Number(delBtn.getAttribute("data-hist-del") || 0));
    }
  });
  document.getElementById("histBody")?.addEventListener("click", (evt) => {
    const viewBtn = evt.target?.closest?.("button[data-hist-view-records]");
    if (!viewBtn) return;
    openServiceRecordsForAsset(
      String(viewBtn.getAttribute("data-hist-view-records") || ""),
      Number(viewBtn.getAttribute("data-hist-asset-id") || 0),
    );
  });
  document.getElementById("loadAssetKpiBtn")?.addEventListener("click", () => loadAssetKpiWeekly());
  document.getElementById("loadReliabilityBtn")?.addEventListener("click", () => loadReliabilityMetrics());
  document.getElementById("exportReliabilityBtn")?.addEventListener("click", () => exportReliabilityToExcel());
  document.getElementById("exportReliabilityExecutiveBtn")?.addEventListener("click", () => exportReliabilityExecutiveToExcel());
  document.getElementById("relCategoryFilter")?.addEventListener("change", () => {
    relRenderAssetOptions(relAssetCatalog);
  });
  document.getElementById("relSelectAllBtn")?.addEventListener("click", () => {
    const sel = document.getElementById("relAssetSelect");
    if (!sel) return;
    Array.from(sel.options).forEach((o) => { o.selected = true; });
  });
  document.getElementById("relMatchAssetKpiScopeBtn")?.addEventListener("click", () => matchReliabilityAssetKpiScope());
  document.getElementById("relClearSelBtn")?.addEventListener("click", () => {
    const sel = document.getElementById("relAssetSelect");
    if (!sel) return;
    Array.from(sel.options).forEach((o) => { o.selected = false; });
  });
  document.getElementById("exportAssetKpiBtn")?.addEventListener("click", () => exportAssetKpiToExcel());
  document.getElementById("exportExecutivePackBtn")?.addEventListener("click", () => exportExecutivePackFromAssetKpi());
  document.getElementById("exportExecutiveKpiPackBtn")?.addEventListener("click", () => exportExecutiveKpiPackBySite());
  document.getElementById("mpRefreshStatusBtn")?.addEventListener("click", () => loadMaintenancePackStatus());
  document.getElementById("mpStatusBody")?.addEventListener("click", (evt) => {
    const gen = evt.target?.closest?.("button[data-mp-gen]");
    if (gen) {
      mpGenerate(String(gen.getAttribute("data-mp-gen") || "weekly"));
      return;
    }
    const open = evt.target?.closest?.("button[data-mp-open]");
    if (open) {
      openMaintenancePackLatest(String(open.getAttribute("data-mp-open") || "weekly"), false);
      return;
    }
    const dl = evt.target?.closest?.("button[data-mp-download]");
    if (dl) {
      openMaintenancePackLatest(String(dl.getAttribute("data-mp-download") || "weekly"), true);
    }
  });
  document.getElementById("akpCategoryFilter")?.addEventListener("change", () => {
    if (akpLastResponse) renderAssetKpiTables(akpLastResponse);
  });
  document.getElementById("akpAvailTarget")?.addEventListener("input", () => {
    if (akpLastResponse) renderAkpKpiChart(akpLastResponse);
  });
  document.getElementById("akpUtilTarget")?.addEventListener("input", () => {
    if (akpLastResponse) renderAkpKpiChart(akpLastResponse);
  });
  document.getElementById("akpChartViewMode")?.addEventListener("change", () => {
    if (akpLastResponse) renderAkpKpiChart(akpLastResponse);
  });
  document.getElementById("akpExportDocxBtn")?.addEventListener("click", exportAssetKpiDocx);
  document.getElementById("akpAssetFilterClear")?.addEventListener("click", () => {
    const sel = document.getElementById("akpAssetFilter");
    if (sel) Array.from(sel.options).forEach((o) => { o.selected = false; });
    loadAssetKpiWeekly();
  });
  document.getElementById("loadWeeklyForumBtn")?.addEventListener("click", () => {
    loadWeeklyForumSummary();
    loadWeeklyForumReviews();
    loadWeeklyForumActions();
  });
  document.getElementById("wfStart")?.addEventListener("change", (evt) => {
    const target = document.getElementById("mpWeekStart");
    if (target) target.value = String(evt.target?.value || "");
    loadWeeklyForumReviews();
  });
  document.getElementById("wfEnd")?.addEventListener("change", (evt) => {
    const target = document.getElementById("mpWeekEnd");
    if (target) target.value = String(evt.target?.value || "");
    loadWeeklyForumReviews();
  });
  document.getElementById("mpWeekStart")?.addEventListener("change", (evt) => {
    const target = document.getElementById("wfStart");
    if (target) target.value = String(evt.target?.value || "");
    loadWeeklyForumReviews();
  });
  document.getElementById("mpWeekEnd")?.addEventListener("change", (evt) => {
    const target = document.getElementById("wfEnd");
    if (target) target.value = String(evt.target?.value || "");
    loadWeeklyForumReviews();
  });
  document.getElementById("wfPresetThisMonth")?.addEventListener("click", () => wfApplyMonthPreset("this"));
  document.getElementById("wfPresetLastMonth")?.addEventListener("click", () => wfApplyMonthPreset("last"));
  document.getElementById("saveRsgProfileBtn")?.addEventListener("click", saveRsgProfile);
  document.getElementById("loadInsightsBtn")?.addEventListener("click", () => loadMaintenanceInsights());
  document.getElementById("openInsightsPdfBtn")?.addEventListener("click", () => openMaintenanceInsightsPdf(false));
  document.getElementById("downloadInsightsPdfBtn")?.addEventListener("click", () => openMaintenanceInsightsPdf(true));
  document.getElementById("downloadInsightsXlsxBtn")?.addEventListener("click", () =>
    openMaintenanceInsightsXlsx().catch((e) => alert(e.message || String(e)))
  );
  document.getElementById("downloadInsightsPartsXlsxBtn")?.addEventListener("click", () =>
    downloadInsightsPartsDemandXlsx().catch((e) => alert(e.message || String(e)))
  );
  document.getElementById("downloadInsightsCostXlsxBtn")?.addEventListener("click", () =>
    downloadInsightsCostPerMachineXlsx().catch((e) => alert(e.message || String(e)))
  );
  document.getElementById("loadGovernanceSignalsBtn")?.addEventListener("click", () => loadGovernanceSignals());
  document.getElementById("reportBuilderRunBtn")?.addEventListener("click", () => runReportBuilderPreview());
  document.getElementById("reportBuilderExportXlsxBtn")?.addEventListener("click", () => exportReportBuilderXlsx());
  document.getElementById("reportBuilderSaveBtn")?.addEventListener("click", () => saveReportBuilderTemplate());
  document.getElementById("reportBuilderDeleteBtn")?.addEventListener("click", () => deleteReportBuilderTemplate());
  document.getElementById("reportBuilderReloadBtn")?.addEventListener("click", () => loadReportBuilderTemplates());
  document.getElementById("reportBuilderDataset")?.addEventListener("change", (evt) => {
    const key = String(evt.target?.value || "");
    const ds = reportBuilderMeta.find((d) => String(d.key || "") === key);
    renderReportBuilderColumns(ds?.columns || [], (ds?.columns || []).slice(0, 4));
  });
  document.getElementById("reportBuilderTemplate")?.addEventListener("change", (evt) => {
    applyReportBuilderTemplate(Number(evt.target?.value || 0));
  });
  document.getElementById("subSaveBtn")?.addEventListener("click", () => saveReportSubscription());
  document.getElementById("subReloadBtn")?.addEventListener("click", () => loadReportSubscriptions());
  document.getElementById("subBody")?.addEventListener("click", (evt) => {
    const editBtn = evt.target?.closest?.("button[data-sub-edit]");
    if (editBtn) {
      const id = Number(editBtn.getAttribute("data-sub-edit") || 0);
      const row = reportSubscriptionsCache.find((x) => Number(x.id || 0) === id);
      if (row) fillSubscriptionForm(row);
      return;
    }
    const sendBtn = evt.target?.closest?.("button[data-sub-send]");
    if (sendBtn) {
      sendSubscriptionNow(Number(sendBtn.getAttribute("data-sub-send") || 0));
      return;
    }
    const delBtn = evt.target?.closest?.("button[data-sub-del]");
    if (delBtn) deleteSubscription(Number(delBtn.getAttribute("data-sub-del") || 0));
  });
  document.getElementById("loadRsgProfilesBtn")?.addEventListener("click", loadRsgProfiles);
  document.getElementById("downloadRsgCsvTemplateBtn")?.addEventListener("click", downloadRsgCsvTemplate);
  document.getElementById("importRsgCsvBtn")?.addEventListener("click", importRsgProfilesCsv);
  document.getElementById("saveWfActionBtn")?.addEventListener("click", saveWeeklyForumAction);
  document.getElementById("loadWfActionsBtn")?.addEventListener("click", loadWeeklyForumActions);
  document.getElementById("saveWfReviewBtn")?.addEventListener("click", saveWeeklyForumReview);
  document.getElementById("loadWfReviewsBtn")?.addEventListener("click", loadWeeklyForumReviews);
  document.getElementById("wfReviewBody")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-wf-review-delete]");
    if (btn) deleteWeeklyForumReview(btn.getAttribute("data-wf-review-delete"));
  });
  document.getElementById("saveWfInputBtn")?.addEventListener("click", saveWeeklyForumInput);
  document.getElementById("loadWfInputsBtn")?.addEventListener("click", loadWeeklyForumInputs);
  document.getElementById("wfAddItemBtn")?.addEventListener("click", addWfDraftItem);
  document.getElementById("ptoSubmitBtn")?.addEventListener("click", () => submitPartsToOrderRequest().catch((e) => {
    const msg = document.getElementById("ptoFormMsg");
    if (msg) {
      msg.className = "message-error";
      msg.textContent = e.message || String(e);
    }
  }));
  document.getElementById("ptoClearBtn")?.addEventListener("click", clearPartsToOrderForm);
  document.getElementById("ptoSelectAllBtn")?.addEventListener("click", selectAllPtoRequests);
  document.getElementById("ptoClearSelBtn")?.addEventListener("click", clearPtoSelection);
  document.getElementById("ptoRefreshBtn")?.addEventListener("click", () => loadPartsToOrderList().catch(() => {}));
  document.getElementById("ptoFilterStatus")?.addEventListener("change", () => loadPartsToOrderList().catch(() => {}));
  document.getElementById("ptoMineOnly")?.addEventListener("change", () => loadPartsToOrderList().catch(() => {}));
  document.getElementById("ptoRfqOpenPdfBtn")?.addEventListener("click", () => openPtoRfqPdf(false));
  document.getElementById("ptoRfqDownloadPdfBtn")?.addEventListener("click", () => openPtoRfqPdf(true));
  document.getElementById("ptoListBody")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-pto-status]");
    if (!btn) return;
    const id = Number(btn.getAttribute("data-pto-status") || 0);
    const status = String(btn.getAttribute("data-pto-next") || "").trim();
    if (!id || !status) return;
    updatePartsToOrderStatus(id, status).catch((e) => alert(e.message || String(e)));
  });
  document.getElementById("mcAddRowBtn")?.addEventListener("click", () => addMcDraftRow());
  document.getElementById("mcAddRowPerTechBtn")?.addEventListener("click", () => addMcDraftRowPerTechnician());
  document.getElementById("mcAdd5RowsBtn")?.addEventListener("click", () => addMcDraftBlankRows(5));
  document.getElementById("mcLoadDayBtn")?.addEventListener("click", loadMcSavedIntoDraft);
  document.getElementById("mcClearDraftBtn")?.addEventListener("click", clearMcDraft);
  document.getElementById("mcSaveAllBtn")?.addEventListener("click", () => saveMcAllRows().catch((e) => {
    const msg = document.getElementById("mcFormMsg");
    if (msg) { msg.className = "message-error"; msg.textContent = e.message || String(e); }
  }));
  document.getElementById("mcRefreshBtn")?.addEventListener("click", () => loadMcDayEntries().catch(() => {}));
  document.getElementById("mcWorkDate")?.addEventListener("change", () => {
    loadMcDayEntries().catch(() => {});
  });
  document.getElementById("mcReportYear")?.addEventListener("change", () => {
    syncMcDateBounds();
    Promise.all([loadMcDayEntries(), loadMcJobCardDatalist()]).catch(() => {});
  });
  document.getElementById("mcSaveRateBtn")?.addEventListener("click", () => saveMcDefaultRate().catch((e) => {
    const msg = document.getElementById("mcFormMsg");
    if (msg) { msg.className = "message-error"; msg.textContent = e.message || String(e); }
  }));
  document.getElementById("mcDownloadXlsxBtn")?.addEventListener("click", () => downloadMcYearXlsx());
  document.getElementById("mcDownloadTimesheetBtn")?.addEventListener("click", () => downloadMcTimesheetXlsx());
  document.getElementById("mcUploadTimesheetBtn")?.addEventListener("click", () => uploadMcTimesheetFile().catch((e) => {
    const msg = document.getElementById("mcTimesheetUploadMsg") || document.getElementById("mcFormMsg");
    if (msg) { msg.className = "message-error"; msg.textContent = e.message || String(e); }
  }));
  document.getElementById("mcTechChips")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-mc-chip-tech]");
    if (!btn) return;
    addMcDraftRow(String(btn.getAttribute("data-mc-chip-tech") || ""));
  });
  document.getElementById("mcDraftBody")?.addEventListener("click", (evt) => {
    const delBtn = evt.target?.closest?.("button[data-mc-draft-del]");
    if (!delBtn) return;
    syncMcDraftFromDom();
    const idx = Number(delBtn.getAttribute("data-mc-draft-del") || -1);
    if (idx < 0 || idx >= mcDraftRows.length) return;
    mcDraftRows.splice(idx, 1);
    refreshMcDraftEditor();
  });
  document.getElementById("mcDraftBody")?.addEventListener("input", (evt) => {
    const fieldEl = evt.target?.closest?.("[data-mc-draft-field]");
    if (!fieldEl) return;
    const idx = Number(fieldEl.getAttribute("data-mc-draft-idx") || -1);
    const field = String(fieldEl.getAttribute("data-mc-draft-field") || "");
    if (idx < 0 || idx >= mcDraftRows.length || !field) return;
    mcDraftRows[idx][field] = String(fieldEl.value || "").trim();
    if (field !== "time_started" && field !== "time_finished") return;
    const row = mcDraftRows[idx];
    const calculated = mcHoursFromTimes(row.time_started, row.time_finished);
    if (calculated == null) return;
    row.hours = String(calculated);
    const hoursEl = document.querySelector(`#mcDraftBody [data-mc-draft-idx="${idx}"][data-mc-draft-field="hours"]`);
    if (hoursEl) hoursEl.value = row.hours;
  });
  document.getElementById("mcDayBody")?.addEventListener("click", (evt) => {
    const addBtn = evt.target?.closest?.("button[data-mc-add-sheet]");
    const delBtn = evt.target?.closest?.("button[data-mc-delete]");
    if (addBtn) {
      const id = Number(addBtn.getAttribute("data-mc-add-sheet") || 0);
      const row = mcDayRowsCache.find((r) => Number(r.id) === id);
      if (row) addMcDraftRowFromSaved(row);
      return;
    }
    if (delBtn) {
      deleteMcEntry(Number(delBtn.getAttribute("data-mc-delete") || 0)).catch((e) => alert(e.message || String(e)));
    }
  });
  document.getElementById("wfInputPlan")?.addEventListener("change", (evt) => {
    const planId = Number(evt.target?.value || 0);
    hydrateWfDraftFromSaved(planId);
  });
  document.getElementById("wfItemsEditorBody")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-wf-item-del]");
    if (!btn) return;
    const idx = Number(btn.getAttribute("data-wf-item-del") || -1);
    if (idx < 0 || idx >= wfDraftItems.length) return;
    wfDraftItems.splice(idx, 1);
    refreshWfDraftEditor();
  });
  document.getElementById("openWeeklyForumPdfBtn")?.addEventListener("click", () => openWeeklyForumPdf(false));
  document.getElementById("downloadWeeklyForumPdfBtn")?.addEventListener("click", () => openWeeklyForumPdf(true));
  document.getElementById("tyrePullLiveHoursBtn")?.addEventListener("click", () => pullTyreLiveHours());
  document.getElementById("tyreAsset")?.addEventListener("change", () => {
    loadTyreLifecycle().catch(() => {});
  });
  document.getElementById("saveTyreInspectionBtn")?.addEventListener("click", () => saveTyreInspection());
  document.getElementById("loadTyreInspectionsBtn")?.addEventListener("click", () => loadTyreInspections());
  document.getElementById("tyreFilterAsset")?.addEventListener("change", () => loadTyreInspections());
  document.getElementById("tyreClearFormBtn")?.addEventListener("click", () => clearTyreForm());
  document.getElementById("tyreSurveyPdfBtn")?.addEventListener("click", () => openTyreSurveyExport("pdf"));
  document.getElementById("tyreSurveyXlsxBtn")?.addEventListener("click", () => openTyreSurveyExport("xlsx"));
  const wiMonth = document.getElementById("wiMonth");
  if (wiMonth && !wiMonth.value) wiMonth.value = new Date().toISOString().slice(0, 7);
  document.getElementById("wiPrevMonthBtn")?.addEventListener("click", () => {
    closeWiDayPanel();
    wiShiftMonth(-1);
    loadWeeklyInspectionCalendar().catch(() => {});
  });
  document.getElementById("wiNextMonthBtn")?.addEventListener("click", () => {
    closeWiDayPanel();
    wiShiftMonth(1);
    loadWeeklyInspectionCalendar().catch(() => {});
  });
  document.getElementById("wiTodayBtn")?.addEventListener("click", () => {
    closeWiDayPanel();
    wiSetMonth(new Date().toISOString().slice(0, 7));
    loadWeeklyInspectionCalendar().catch(() => {});
  });
  document.getElementById("wiRefreshBtn")?.addEventListener("click", () => loadWeeklyInspectionCalendar());
  document.getElementById("wiPrintBtn")?.addEventListener("click", () => printWeeklyInspectionCalendar());
  document.getElementById("wiOpenPdfBtn")?.addEventListener("click", () => openWeeklyInspectionPdf(false));
  document.getElementById("wiToggleComplianceBtn")?.addEventListener("click", () => toggleWiCompliancePanel());
  document.getElementById("wiAddEquipmentBtn")?.addEventListener("click", () => {
    document.getElementById("wiRosterSection")?.scrollIntoView({ behavior: "smooth", block: "start" });
    document.getElementById("wiAssetSelect")?.focus();
  });
  document.getElementById("wiAddAssetBtn")?.addEventListener("click", () => addWeeklyInspectionAsset());
  document.getElementById("wiClearRosterBtn")?.addEventListener("click", () => clearWeeklyInspectionRoster());
  document.getElementById("wiAddSlotBtn")?.addEventListener("click", () => addWeeklyInspectionSlot());
  document.getElementById("wiCopyDayBtn")?.addEventListener("click", () => copyWeeklyInspectionDay());
  document.getElementById("wiCloseDayPanelBtn")?.addEventListener("click", () => closeWiDayPanel());
  document.getElementById("wiMonth")?.addEventListener("change", () => {
    closeWiDayPanel();
    loadWeeklyInspectionCalendar();
  });
  document.getElementById("wiCalendarWrap")?.addEventListener("click", (evt) => {
    const slotBtn = evt.target?.closest?.("button[data-wi-slot-status]");
    if (slotBtn) {
      evt.stopPropagation();
      cycleWeeklyInspectionSlot(
        Number(slotBtn.getAttribute("data-wi-slot-status") || 0),
        String(slotBtn.getAttribute("data-wi-status") || "pending"),
      );
      return;
    }
    const dayEl = evt.target?.closest?.("[data-wi-open-day]");
    if (!dayEl) return;
    openWiDayPanel(String(dayEl.getAttribute("data-wi-open-day") || ""));
  });
  document.getElementById("wiUpcomingList")?.addEventListener("click", (evt) => {
    const dayBtn = evt.target?.closest?.("button[data-wi-open-day]");
    if (!dayBtn) return;
    const date = String(dayBtn.getAttribute("data-wi-open-day") || "");
    const month = date.slice(0, 7);
    if (month && month !== wiCurrentMonth()) {
      wiSetMonth(month);
      loadWeeklyInspectionCalendar().then(() => openWiDayPanel(date)).catch(() => {});
      return;
    }
    openWiDayPanel(date);
  });
  document.getElementById("wiDayPanel")?.addEventListener("click", (evt) => {
    const slotBtn = evt.target?.closest?.("button[data-wi-slot-status]");
    if (slotBtn) {
      cycleWeeklyInspectionSlot(
        Number(slotBtn.getAttribute("data-wi-slot-status") || 0),
        String(slotBtn.getAttribute("data-wi-status") || "pending"),
      );
      return;
    }
    const removeBtn = evt.target?.closest?.("button[data-wi-slot-remove]");
    if (!removeBtn) return;
    removeWeeklyInspectionSlot(Number(removeBtn.getAttribute("data-wi-slot-remove") || 0));
  });
  document.getElementById("wiAssetList")?.addEventListener("click", (evt) => {
    const saveBtn = evt.target?.closest?.("button[data-wi-save-est]");
    if (saveBtn) {
      saveWeeklyInspectionAssetEst(Number(saveBtn.getAttribute("data-wi-save-est") || 0));
      return;
    }
    const btn = evt.target?.closest?.("button[data-wi-remove]");
    if (!btn) return;
    removeWeeklyInspectionAsset(Number(btn.getAttribute("data-wi-remove") || 0));
  });
  initTyreLayout();
  const tyreDate = document.getElementById("tyreInspectionDate");
  if (tyreDate && !tyreDate.value) tyreDate.value = new Date().toISOString().slice(0, 10);
  const tyreSurveyMonthEl = document.getElementById("tyreSurveyMonth");
  if (tyreSurveyMonthEl && !tyreSurveyMonthEl.value) tyreSurveyMonthEl.value = new Date().toISOString().slice(0, 7);
  loadTyreInspections();
  if (typeof window.bindUndercarriageEvents === "function") window.bindUndercarriageEvents();
  if (typeof window.bindTyreQrAdmin === "function") window.bindTyreQrAdmin();
  document.getElementById("wfActionBody")?.addEventListener("change", (evt) => {
    const sel = evt.target?.closest?.("select[data-wf-action-status]");
    if (!sel) return;
    const id = Number(sel.getAttribute("data-wf-action-status") || 0);
    const status = String(sel.value || "open");
    updateWeeklyForumActionStatus(id, status)
      .then(() => loadWeeklyForumActions())
      .catch((e) => alert(`Failed to update status: ${e.message || e}`));
  });
  document.getElementById("rsgProfilesBody")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-rsg-edit]");
    if (!btn) return;
    fillRsgProfileForm(btn.getAttribute("data-rsg-edit") || "");
  });
  document.getElementById("insightsPredictive")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-insights-asset-code]");
    if (!btn) return;
    const code = String(btn.getAttribute("data-insights-asset-code") || "").trim().toUpperCase();
    if (!code) return;
    const sel = document.getElementById("miFilterAsset");
    if (sel) {
      const opt = Array.from(sel.options || []).find((o) => String(o.text || "").toUpperCase().includes(code));
      if (opt) sel.value = String(opt.value || "");
    }
    setTopView("mi");
    loadManagerInspections().catch(() => {});
  });
  loadMaintenanceInsights().catch(() => {});
  loadGovernanceSignals().catch(() => {});
  loadReportBuilderMeta().then(() => loadReportBuilderTemplates()).catch(() => {});
  loadReportSubscriptions().catch(() => {});
  scrollToSection("maintenance");
  loadHistogramEvents().catch(() => {});
});

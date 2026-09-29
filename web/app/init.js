// IRONLOG/web/app/init.js — App start-up and event wiring.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function init() {
  await disableLegacyServiceWorkers();
  await loadAuthConfig();
  await tryInitialSession();
  initSidebar();
  initDashboardActionHub();
  initTabs();
  initWorkshopLibraryTab();
  initIronmindHelpUi();
  initSectionCollapseToggles();
  initSessionControls();
  initVehicleCheckTab();
  initSettingsDropdown();
  initGlobalSearch();
  initReportCardCollapsible();
  document.querySelectorAll("[data-utility-hub]").forEach((hub) => {
    hub.addEventListener("click", (evt) => {
      const btn = evt.target?.closest?.("button[data-open-card]");
      if (!btn) return;
      openUtilityWorkflow(btn.dataset.openCard);
    });
  });
  initReportsHub();
  initTasks();
  initTelematicsFaultBanner();
  initCartrackSpeedFloat();
  applyRoleVisibility();
  if (!isBareChildTabEmbed()) {
    resolveInitialTabFromUrl();
  }
  applyBareChildTabView();
  refreshMyWorkIfVisible();
  applyI18n();
  applyGlobalPageTranslation();

  const dateEl = qs("date");
  if (dateEl) dateEl.value = todayLocalYmd();
  const mtdMonthEl = qs("mtdOpeningMonth");
  if (mtdMonthEl && !mtdMonthEl.value) {
    mtdMonthEl.value = (dateEl?.value || todayLocalYmd()).slice(0, 7);
  }

  qs("refresh")?.addEventListener("click", () =>
    loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message))
  );
  qs("kpiDebugToggle")?.addEventListener("change", () =>
    loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message))
  );
  qs("loadReliability")?.addEventListener("click", () =>
    loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message))
  );
  qs("ironmindRefreshBtn")?.addEventListener("click", () =>
    refreshIronmindInsight().catch((e) => setStatus("BORRIS refresh error: " + e.message))
  );
  setInterval(() => {
    loadIronmindHealth().catch(() => {});
  }, 30000);
  qs("ironmindSaveSettingsBtn")?.addEventListener("click", () =>
    saveIronmindSettings().catch((e) => setStatus("Borris settings error: " + (e.message || e)))
  );
  qs("ironmindWeeklyPlanBtn")?.addEventListener("click", planIronmindWeek);
  qs("ironmindAskBtn")?.addEventListener("click", () =>
    askIronmindQuestion().catch((e) => setStatus("BORRIS ask error: " + e.message))
  );
  qs("ironmindResetMemoryBtn")?.addEventListener("click", () =>
    resetIronmindAskMemory().catch((e) => setStatus("BORRIS reset memory error: " + e.message))
  );
  qs("ironmindAskInput")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      askIronmindQuestion().catch((err) => setStatus("BORRIS ask error: " + err.message));
    }
  });
  hydrateIronmindAskMemory().catch(() => {});
  qs("saveThresholds")?.addEventListener("click", () => saveThresholdsFromUI());
  qs("saveLdvThresholds")?.addEventListener("click", () => saveLdvPrestartThresholdsFromUI());
  qs("ironmindSummary")?.addEventListener("click", (e) => {
    const el = e.target instanceof HTMLElement ? e.target : null;
    if (!el) return;
    const drillKey = el.dataset.ironmindDrill;
    const assetCode = el.dataset.ironmindAsset;
    if (drillKey) ironmindDrillDown(drillKey);
    if (assetCode) ironmindGoToAsset(assetCode).catch(() => {});
  });
  qs("ironmindHistoryList")?.addEventListener("click", (e) => {
    const el = e.target instanceof HTMLElement ? e.target.closest("button[data-ironmind-history-id]") : null;
    if (!el) return;
    const rowEl = el.closest(".item");
    if (!rowEl?.dataset?.ironmindRow) return;
    try {
      const row = JSON.parse(rowEl.dataset.ironmindRow);
      renderIronmindReport(row);
      setStatus(`Opened BORRIS report for ${row?.report_date || "-"}.`);
    } catch (_) {
      setStatus("Unable to open selected BORRIS report.");
    }
  });
  qs("riskBoardList")?.addEventListener("click", (e) => {
    const target = e.target instanceof HTMLElement ? e.target : null;
    if (!target) return;
    const openBtn = target.closest("button[data-ironmind-risk-asset]");
    if (openBtn) {
      const code = String(openBtn.getAttribute("data-ironmind-risk-asset") || "").trim();
      if (!code) return;
      ironmindGoToAsset(code).catch(() => {});
      return;
    }
    const woBtn = target.closest("button[data-ironmind-risk-wo]");
    if (woBtn) {
      const code = String(woBtn.getAttribute("data-ironmind-risk-wo") || "").trim();
      if (!code) return;
      (async () => {
        const downDesc = "BORRIS predicted risk work order";
        const res = await fetchJson(`${API}/api/breakdowns/ensure-open`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            asset_code: code,
            breakdown_date: date,
            description: downDesc,
            critical: false,
          }),
        });
        const woId = Number(res?.primary_work_order_id || 0);
        setStatus(woId > 0 ? `WO #${woId} ready for ${code}.` : `Work order ensured for ${code}.`);
        await loadDashboard().catch(() => {});
      })().catch((err) => setStatus("Create WO failed: " + (err.message || err)));
    }
  });
  qs("ironmindShowMissingDays")?.addEventListener("change", () => {
    loadIronmindHistory({ silent: true }).catch(() => {});
  });
  qs("ironmindReloadHistory")?.addEventListener("click", () => {
    loadIronmindHistory().catch((e) => setStatus("BORRIS history error: " + e.message));
  });
  qs("ironmindRsgPlanBtn")?.addEventListener("click", () => {
    generateIronmindRsgPlan(false).catch((e) => setStatus("RSG plan error: " + (e.message || e)));
  });
  qs("ironmindRsgPreviewPdfBtn")?.addEventListener("click", () => {
    previewIronmindRsgPdf().catch((e) => setStatus("RSG preview error: " + (e.message || e)));
  });
  qs("ironmindRsgDownloadPdfBtn")?.addEventListener("click", () => {
    downloadIronmindRsgPdf().catch((e) => setStatus("RSG download error: " + (e.message || e)));
  });
  qs("ironmindRsgCreateWoBtn")?.addEventListener("click", () => {
    generateIronmindRsgPlan(true).catch((e) => setStatus("RSG create WO error: " + (e.message || e)));
  });
  qs("saveDocHeaderBtn")?.addEventListener("click", () =>
    saveDocHeader().catch((e) => setStatus("Header save error: " + e.message))
  );
  qs("loadDocHeadersBtn")?.addEventListener("click", () =>
    loadDocHeaders().catch((e) => setStatus("Header load error: " + e.message))
  );
  qs("generateDocDraftBtn")?.addEventListener("click", () =>
    generateDocDraft().catch((e) => setStatus("Draft generate error: " + e.message))
  );
  qs("generateDocDraftFromRequestBtn")?.addEventListener("click", () =>
    generateDocDraftFromRequest().catch((e) => setStatus("Draft request generate error: " + e.message))
  );
  qs("aiSmartRunBtn")?.addEventListener("click", () =>
    runAiSmart().catch((e) => setStatus("Smart AI error: " + e.message))
  );
  qs("askJakesBtn")?.addEventListener("click", () =>
    askJakes().catch((e) => setStatus("Ask Jakes error: " + e.message))
  );
  qs("askJakesPresetHydraulics")?.addEventListener("click", () =>
    applyAskJakesPreset("hydraulics")
  );
  qs("askJakesPresetStarting")?.addEventListener("click", () =>
    applyAskJakesPreset("starting")
  );
  qs("askJakesPresetOverheat")?.addEventListener("click", () =>
    applyAskJakesPreset("overheat")
  );
  qs("askJakesUseAsNotesBtn")?.addEventListener("click", () =>
    useAskJakesAnswerAsNotes()
  );
  qs("speakDocDraftBtn")?.addEventListener("click", () =>
    speakDocDraft()
  );
  qs("stopSpeakDocDraftBtn")?.addEventListener("click", () =>
    stopSpeakingDocDraft()
  );
  qs("docApproveYesBtn")?.addEventListener("click", () =>
    decideDocDraft(true).catch((e) => setStatus("Draft decision error: " + e.message))
  );
  qs("docApproveNoBtn")?.addEventListener("click", () =>
    decideDocDraft(false).catch((e) => setStatus("Draft decision error: " + e.message))
  );
  qs("openDocDraftPdfBtn")?.addEventListener("click", () =>
    openDocDraftPdf(false)
  );
  qs("downloadDocDraftPdfBtn")?.addEventListener("click", () =>
    openDocDraftPdf(true)
  );
  qs("openDocDraftWordBtn")?.addEventListener("click", () =>
    openDocDraftWord(false)
  );
  qs("downloadDocDraftWordBtn")?.addEventListener("click", () =>
    openDocDraftWord(true)
  );
  qs("openDocRegisterPdfBtn")?.addEventListener("click", () =>
    openDocRegisterPdf(false)
  );
  qs("downloadDocRegisterPdfBtn")?.addEventListener("click", () =>
    openDocRegisterPdf(true)
  );
  qs("openDocRegisterWordBtn")?.addEventListener("click", () =>
    openDocRegisterWord(false)
  );
  qs("downloadDocRegisterWordBtn")?.addEventListener("click", () =>
    openDocRegisterWord(true)
  );
  qs("loadDocDraftsBtn")?.addEventListener("click", () =>
    loadDocDrafts().catch((e) => setStatus("Draft list error: " + e.message))
  );
  qs("docDraftsCurrentOnly")?.addEventListener("change", () =>
    loadDocDrafts().catch((e) => setStatus("Draft list error: " + e.message))
  );
  qs("loadLube")?.addEventListener("click", () =>
    loadLubeUsage().catch((e) => setStatus("Lube error: " + e.message))
  );
  qs("lubeFilterAsset")?.addEventListener("input", () => {
    if (lubeUsageCache) renderLubeUsageTable(lubeUsageCache);
  });
  qs("lubeFilterOilType")?.addEventListener("input", () => {
    if (lubeUsageCache) renderLubeUsageTable(lubeUsageCache);
  });
  qs("lubeHidePartLike")?.addEventListener("change", () => {
    if (lubeUsageCache) renderLubeUsageTable(lubeUsageCache);
  });
  qs("loadLubeAnalytics")?.addEventListener("click", () =>
    loadLubeAnalytics().catch((e) => setStatus("Lube analytics error: " + e.message))
  );
  qs("createRequisition")?.addEventListener("click", () =>
    createRequisition().catch((e) => setStatus("Requisition create error: " + e.message))
  );
  qs("prLoadPoList")?.addEventListener("click", () =>
    loadPurchaseOrders().catch((e) => setStatus("PO load error: " + e.message))
  );
  qs("prPoStatusFilter")?.addEventListener("change", () =>
    loadPurchaseOrders().catch((e) => setStatus("PO load error: " + e.message))
  );
  qs("prPostReceiptBtn")?.addEventListener("click", () =>
    postPoReceipt().catch((e) => setStatus("PO receipt error: " + e.message))
  );
  qs("prCaptureInvoiceBtn")?.addEventListener("click", () =>
    capturePoInvoice().catch((e) => setStatus("Invoice capture error: " + e.message))
  );
  qs("prRunMatchBtn")?.addEventListener("click", () =>
    runPoThreeWayMatch().catch((e) => setStatus("3-way match error: " + e.message))
  );
  qs("prLoadExceptionsBtn")?.addEventListener("click", () =>
    loadProcurementExceptions().catch((e) => setStatus("Exception load error: " + e.message))
  );
  qs("prExceptionStatus")?.addEventListener("change", () =>
    loadProcurementExceptions().catch((e) => setStatus("Exception load error: " + e.message))
  );
  qs("prBuildJournalBtn")?.addEventListener("click", () =>
    buildProcurementJournals().catch((e) => setStatus("Journal build error: " + e.message))
  );
  qs("prExportJournalCsvBtn")?.addEventListener("click", () => {
    try {
      exportProcurementJournalsCsv();
    } catch (e) {
      setStatus("Journal CSV export error: " + (e.message || e));
    }
  });
  qs("prExportJournalXlsxBtn")?.addEventListener("click", () => {
    try {
      exportProcurementJournalsXlsx();
    } catch (e) {
      setStatus("Journal XLSX export error: " + (e.message || e));
    }
  });
  qs("loadRequisitions")?.addEventListener("click", () =>
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message))
  );
  qs("prStatusFilter")?.addEventListener("change", () => {
    setProcurementKpiFilter("all");
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prTierFilter")?.addEventListener("change", () => {
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prKpiAll")?.addEventListener("click", () => {
    setProcurementKpiFilter("all");
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prKpiApprovedOpen")?.addEventListener("click", () => {
    const statusEl = qs("prStatusFilter");
    if (statusEl) statusEl.value = "";
    setProcurementKpiFilter("approved_open");
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prKpiInFlow")?.addEventListener("click", () => {
    const statusEl = qs("prStatusFilter");
    if (statusEl) statusEl.value = "";
    setProcurementKpiFilter("in_flow");
    loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
  });
  qs("prSaveChainConfig")?.addEventListener("click", () => {
    try {
      saveProcurementChainConfig();
      updateProcurementChainPreview();
    } catch (e) {
      setStatus(`Save chain rules failed: ${e.message || e}`);
    }
  });
  ["prValue", "prTier1Max", "prTier1Chain", "prTier2Max", "prTier2Chain", "prTier3Chain", "prApproverChain"].forEach((id) => {
    qs(id)?.addEventListener("input", updateProcurementChainPreview);
  });
  qs("loadLubeMaps")?.addEventListener("click", () =>
    loadLubeMappings().catch((e) => setStatus("Lube mapping load error: " + e.message))
  );
  qs("saveLubeMap")?.addEventListener("click", () =>
    saveLubeMapping().catch((e) => setStatus("Lube mapping save error: " + e.message))
  );
  qs("saveFuelLog")?.addEventListener("click", () =>
    saveFuelLog().catch((e) => setStatus("Fuel log error: " + e.message))
  );
  qs("fuelMassImportBtn")?.addEventListener("click", () =>
    importFuelMassPaste().catch((e) => setStatus("Fuel mass import error: " + e.message))
  );
  qs("fuelAsset")?.addEventListener("change", () => {
    applyAssetCostCenterToInputs(qs("fuelAsset")?.value);
    syncFuelUnitFromAsset(qs("fuelAsset")?.value, "input").catch(() => {});
  });
  qs("mlAsset")?.addEventListener("change", () => {
    applyAssetCostCenterToInputs(qs("mlAsset")?.value);
  });
  qs("fuelMeterUnit")?.addEventListener("change", () => {
    const mode = String(qs("fuelMeterUnit")?.value || "hours").toLowerCase() === "km" ? "km" : "hours";
    const meterInput = qs("fuelHoursRun");
    if (meterInput) meterInput.placeholder = mode === "km" ? "Distance since fill (km)" : "Hours since fill";
  });
  qs("loadFuelBaseline")?.addEventListener("click", () =>
    loadFuelBaseline().catch((e) => setStatus("Fuel baseline error: " + e.message))
  );
  qs("fuelBaseAsset")?.addEventListener("change", () => {
    syncFuelUnitFromAsset(qs("fuelBaseAsset")?.value, "both").catch(() => {});
  });
  qs("saveFuelBaseline")?.addEventListener("click", () =>
    saveFuelBaseline().catch((e) => setStatus("Fuel baseline error: " + e.message))
  );
  qs("createDispatchTrip")?.addEventListener("click", () =>
    createDispatchTrip().catch((e) => setStatus("Dispatch create error: " + e.message))
  );
  qs("saveDispatchPod")?.addEventListener("click", () =>
    saveDispatchPod().catch((e) => setStatus("Dispatch POD error: " + e.message))
  );
  qs("createDispatchException")?.addEventListener("click", () =>
    createDispatchException().catch((e) => setStatus("Dispatch exception error: " + e.message))
  );
  qs("loadDispatchExceptions")?.addEventListener("click", () =>
    loadDispatchExceptions().catch((e) => setStatus("Dispatch exceptions load error: " + e.message))
  );
  qs("loadDispatchTrips")?.addEventListener("click", () =>
    loadDispatchTrips().catch((e) => setStatus("Dispatch load error: " + e.message))
  );
  qs("dpStatusFilter")?.addEventListener("change", () =>
    loadDispatchTrips().catch((e) => setStatus("Dispatch load error: " + e.message))
  );
  qs("dpVarTolerance")?.addEventListener("change", () =>
    loadDispatchTrips().catch((e) => setStatus("Dispatch load error: " + e.message))
  );
  qs("dpOnlyBreaches")?.addEventListener("change", () =>
    loadDispatchTrips().catch((e) => setStatus("Dispatch load error: " + e.message))
  );
  qs("dpExStatusFilter")?.addEventListener("change", () =>
    loadDispatchExceptions().catch((e) => setStatus("Dispatch exceptions load error: " + e.message))
  );
  qs("dpExOnlyOpen")?.addEventListener("change", () =>
    loadDispatchExceptions().catch((e) => setStatus("Dispatch exceptions load error: " + e.message))
  );
  qs("loadQualityCenter")?.addEventListener("click", () =>
    loadQualityCenter().catch((e) => setStatus("Quality center load error: " + e.message))
  );
  qs("qSeverityFilter")?.addEventListener("change", () =>
    loadQualityCenter().catch((e) => setStatus("Quality center load error: " + e.message))
  );
  qs("qTypeFilter")?.addEventListener("change", () =>
    loadQualityCenter().catch((e) => setStatus("Quality center load error: " + e.message))
  );
  qs("saveOperationEntry")?.addEventListener("click", () =>
    saveOperationEntry().catch((e) => setStatus("Operations save error: " + e.message))
  );
  qs("saveSiteDailyEntry")?.addEventListener("click", () =>
    saveSiteDailyEntry().catch((e) => setStatus("Site daily save error: " + e.message))
  );
  qs("loadSiteDailyEntries")?.addEventListener("click", () =>
    loadSiteDailyEntries().catch((e) => setStatus("Site daily load error: " + e.message))
  );
  qs("saveSiteEquipmentUsage")?.addEventListener("click", () =>
    saveSiteEquipmentUsage().catch((e) => setStatus("Site equipment link error: " + e.message))
  );
  qs("loadSiteEquipmentUsage")?.addEventListener("click", () =>
    loadSiteEquipmentUsage().catch((e) => setStatus("Site equipment load error: " + e.message))
  );
  qs("saveSiteTarget")?.addEventListener("click", () =>
    saveSiteTarget().catch((e) => setStatus("Site target save error: " + e.message))
  );
  qs("loadSiteTargets")?.addEventListener("click", () =>
    loadSiteTargets().catch((e) => setStatus("Site target load error: " + e.message))
  );
  qs("saveSiteDelay")?.addEventListener("click", () =>
    saveSiteDelay().catch((e) => setStatus("Site delay save error: " + e.message))
  );
  qs("loadSiteDelays")?.addEventListener("click", () =>
    loadSiteDelays().catch((e) => setStatus("Site delay load error: " + e.message))
  );
  qs("saveSiteZone")?.addEventListener("click", () =>
    saveSiteZone().catch((e) => setStatus("Site zone save error: " + e.message))
  );
  qs("loadSiteZones")?.addEventListener("click", () =>
    loadSiteZones().catch((e) => setStatus("Site zone load error: " + e.message))
  );
  qs("loadSiteDashboard")?.addEventListener("click", () =>
    loadSiteDashboard().catch((e) => setStatus("Site dashboard load error: " + e.message))
  );
  qs("saveOperationsClosingDraft")?.addEventListener("click", () =>
    saveOperationsClosing(false).catch((e) => setStatus("Operations closing error: " + e.message))
  );
  qs("closeOperationsDay")?.addEventListener("click", () =>
    saveOperationsClosing(true).catch((e) => setStatus("Operations close day error: " + e.message))
  );
  qs("reopenOperationsDay")?.addEventListener("click", () =>
    reopenOperationsDay().catch((e) => setStatus("Operations reopen day error: " + e.message))
  );
  qs("opDate")?.addEventListener("change", () =>
    loadOperationsClosingForDate((qs("opDate")?.value || "").trim()).catch((e) => setStatus("Operations closing load error: " + e.message))
  );
  qs("loadOperations")?.addEventListener("click", () =>
    loadOperations().catch((e) => setStatus("Operations load error: " + e.message))
  );
  qs("opClientMetric")?.addEventListener("change", () =>
    loadOperations().catch((e) => setStatus("Operations load error: " + e.message))
  );
  qs("opClientTopN")?.addEventListener("change", () =>
    loadOperations().catch((e) => setStatus("Operations load error: " + e.message))
  );
  qs("loadCostSettings")?.addEventListener("click", () =>
    loadCostSettings().catch((e) => setStatus("Cost settings error: " + e.message))
  );
  qs("saveCostSettings")?.addEventListener("click", () =>
    saveCostSettings().catch((e) => setStatus("Cost settings error: " + e.message))
  );
  qs("saveCostAssetRates")?.addEventListener("click", () =>
    saveCostAssetRates().catch((e) => setStatus("Cost asset rates error: " + e.message))
  );
  qs("saveCostPartRate")?.addEventListener("click", () =>
    saveCostPartRate().catch((e) => setStatus("Cost part rate error: " + e.message))
  );
  qs("loadFuelBenchmark")?.addEventListener("click", () =>
    loadFuelBenchmark().catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("loadShiftScenario")?.addEventListener("click", () =>
    loadShiftScenario().catch((e) => setStatus("Shift scenario error: " + e.message))
  );
  qs("downloadShiftScenarioXlsx")?.addEventListener("click", () =>
    downloadShiftScenarioXlsx().catch((e) => setStatus("Shift scenario export error: " + e.message))
  );
  qs("fuelPresetQ1")?.addEventListener("click", () =>
    applyFuelPeriodPreset("q1").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelPresetQ2")?.addEventListener("click", () =>
    applyFuelPeriodPreset("q2").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelPresetQ3")?.addEventListener("click", () =>
    applyFuelPeriodPreset("q3").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelPresetYtd")?.addEventListener("click", () =>
    applyFuelPeriodPreset("ytd").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelPresetMtd")?.addEventListener("click", () =>
    applyFuelPeriodPreset("mtd").catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelDupOnly")?.addEventListener("change", () =>
    loadFuelBenchmark().catch((e) => setStatus("Fuel benchmark error: " + e.message))
  );
  qs("fuelEquipSelectAll")?.addEventListener("click", () => {
    const host = qs("fuelEquipFilterList");
    host?.querySelectorAll('input[type="checkbox"][data-fuel-equip]')?.forEach((b) => { b.checked = true; });
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("fuelEquipClear")?.addEventListener("click", () => {
    const host = qs("fuelEquipFilterList");
    host?.querySelectorAll('input[type="checkbox"][data-fuel-equip]')?.forEach((b) => { b.checked = false; });
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("fuelEquipTypeFilter")?.addEventListener("change", (evt) => {
    const type = String(evt.target?.value || "").trim();
    if (!type) return;
    const host = qs("fuelEquipFilterList");
    if (!host) return;
    // Uncheck all, then check only the matching type
    host.querySelectorAll('input[type="checkbox"][data-fuel-equip]').forEach((b) => { b.checked = false; });
    host.querySelectorAll(`label[data-fuel-type="${CSS.escape(type)}"] input[type="checkbox"]`).forEach((b) => { b.checked = true; });
    // Reset the select so it can be used again next time
    evt.target.value = "";
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("fuelEquipViewMode")?.addEventListener("change", () => {
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("fuelEquipFilterList")?.addEventListener("change", (evt) => {
    const t = evt.target;
    if (!(t instanceof HTMLInputElement) || t.type !== "checkbox") return;
    if (!t.hasAttribute("data-fuel-equip")) return;
    renderFuelEquipmentChart(window.__fuelBenchmarkChartRows || []);
  });
  qs("loadFuelSnapshots")?.addEventListener("click", () =>
    loadFuelSnapshots().catch((e) => setStatus("Fuel snapshots error: " + e.message))
  );
  qs("fuelBenchmarkList")?.addEventListener("click", (evt) => {
    const pdfBtn = evt.target?.closest?.("button[data-fuel-machine-pdf]");
    if (pdfBtn) {
      const code = String(pdfBtn.getAttribute("data-fuel-machine-pdf") || "").trim();
      if (!code) return;
      openFuelMachineHistoryPdf(code, false);
      return;
    }

    const saveBtn = evt.target?.closest?.("button[data-fuel-save]");
    if (saveBtn) {
      const rowEl = saveBtn.closest(".item");
      const mountEl = rowEl?.querySelector?.(".fuel-inline-history");
      const code = String(mountEl?.getAttribute?.("data-code") || "");
      saveFuelMachineHoursInline(saveBtn)
        .then(() => Promise.all([
          loadFuelBenchmark().catch(() => {}),
          code && mountEl ? loadFuelMachineDailyInline(code, mountEl).catch(() => {}) : Promise.resolve(),
        ]))
        .then(() => setStatus("Machine hours updated."))
        .catch((e) => setStatus("Machine hours update failed: " + (e.message || e)));
      return;
    }

    const delBtn = evt.target?.closest?.("button[data-fuel-delete]");
    if (delBtn) {
      const logId = Number(delBtn.getAttribute("data-fuel-delete") || 0);
      const rowEl = delBtn.closest(".item");
      const mountEl = rowEl?.querySelector?.(".fuel-inline-history");
      const code = String(mountEl?.getAttribute?.("data-code") || "");
      deleteFuelLogEntry(logId)
        .then(() => Promise.all([
          loadFuelBenchmark().catch(() => {}),
          code && mountEl ? loadFuelMachineDailyInline(code, mountEl).catch(() => {}) : Promise.resolve(),
        ]))
        .then(() => setStatus("Fuel input deleted."))
        .catch((e) => setStatus("Delete fuel input failed: " + (e.message || e)));
      return;
    }

    const btn = evt.target?.closest?.("button[data-fuel-machine]");
    if (!btn) return;
    const code = String(btn.getAttribute("data-fuel-machine") || "").trim();
    if (!code) return;
    const rowEl = btn.closest(".item");
    const mountEl = rowEl?.querySelector?.(".fuel-inline-history");
    if (!mountEl) return;
    const opened = mountEl.getAttribute("data-opened") === "1";
    const openedCode = String(mountEl.getAttribute("data-code") || "");
    if (opened && openedCode === code) {
      mountEl.innerHTML = "";
      mountEl.setAttribute("data-opened", "0");
      mountEl.setAttribute("data-code", "");
      return;
    }
    mountEl.setAttribute("data-opened", "1");
    mountEl.setAttribute("data-code", code);
    loadFuelMachineDailyInline(code, mountEl).catch((e) => {
      mountEl.innerHTML = `<small>Machine history error: ${String(e.message || e)}</small>`;
      setStatus("Machine fuel consumption error: " + (e.message || e));
    });
  });
  qs("openFuelBenchmarkPdf")?.addEventListener("click", () => openFuelBenchmarkPdf(false));
  qs("downloadFuelBenchmarkPdf")?.addEventListener("click", () => openFuelBenchmarkPdf(true));
  qs("downloadFuelBenchmarkXlsx")?.addEventListener("click", () => openFuelBenchmarkXlsx());
  qs("openFuelReconPdf")?.addEventListener("click", () => openFuelReconciliationPdf(false));
  qs("downloadFuelReconPdf")?.addEventListener("click", () => openFuelReconciliationPdf(true));
  qs("downloadFuelReconXlsx")?.addEventListener("click", () => openFuelReconciliationXlsx());
  qs("runFuelReconYtd")?.addEventListener("click", () =>
    runFuelReconciliation(true).catch((e) => setStatus("Fuel reconciliation error: " + e.message))
  );
  qs("runFuelReconRange")?.addEventListener("click", () =>
    runFuelReconciliation(false).catch((e) => setStatus("Fuel reconciliation error: " + e.message))
  );
  qs("downloadExecutivePackXlsx")?.addEventListener("click", () => {
    downloadExecutivePackExcel().catch((e) => setStatus("Executive pack error: " + e.message));
  });
  qs("loadStockMonitor")?.addEventListener("click", () =>
    loadStockMonitor().catch((e) => setStatus("Stock monitor error: " + e.message))
  );
  qs("spLoad")?.addEventListener("click", () =>
    loadStockOnHandPage().catch((e) => setStatus("Stock on hand error: " + e.message))
  );
  ensureGmStockReportDate();
  updateGmStockReportHelp();
  qs("gmStockReportPeriod")?.addEventListener("change", updateGmStockReportHelp);
  qs("downloadGmStockReportXlsx")?.addEventListener("click", () => downloadGmStockReportXlsx());
  qs("spSort")?.addEventListener("change", () => refreshStockInventoryDisplay());
  qs("spOnlyLow")?.addEventListener("change", () => refreshStockInventoryDisplay());
  qs("spList")?.addEventListener("click", (evt) => {
    const btn = evt.target?.closest?.("button[data-stock-action]");
    if (!btn) return;
    openStockAction(btn.closest("[data-stock-code]"), btn.getAttribute("data-stock-action"));
  });
  qs("spFilter")?.addEventListener("input", () => {
    if (stockPageData.rows.length) refreshStockInventoryDisplay();
  });
  qs("spExportCsv")?.addEventListener("click", exportStockOnHandCsv);
  qs("spOpenPdf")?.addEventListener("click", openStockOnHandPdf);
  qs("smrLoad")?.addEventListener("click", () =>
    loadStockMovementsReport().catch((e) => setStatus("Stock movements report error: " + e.message))
  );
  qs("smrExportCsv")?.addEventListener("click", exportStockMovementsReportCsv);
  qs("smrOpenPdf")?.addEventListener("click", openStockMovementsReportPdf);
  qs("spoLoad")?.addEventListener("click", () =>
    loadStoresPartOrders().catch((e) => setStatus("Parts purchases error: " + e.message))
  );
  qs("spoSave")?.addEventListener("click", () =>
    saveStoresPartOrder().catch((e) => {
      const msg = qs("spoFormMsg");
      if (msg) msg.textContent = e.message || String(e);
    })
  );
  qs("spoClear")?.addEventListener("click", clearStoresPartOrderForm);
  qs("spoFilterStatus")?.addEventListener("change", () => {
    loadStoresPartOrders().catch(() => {});
  });
  qs("spoList")?.addEventListener("change", (evt) => {
    const sel = evt.target?.closest?.("select[data-spo-status]");
    if (!sel) return;
    const id = Number(sel.getAttribute("data-spo-status") || 0);
    const status = String(sel.value || "").trim();
    if (!id) return;
    updateStoresPartOrderStatus(id, status).catch((e) => alert(e.message || String(e)));
  });
  qs("spoList")?.addEventListener("click", (evt) => {
    const saveBtn = evt.target?.closest?.("button[data-spo-save]");
    if (saveBtn) {
      saveStoresPartOrderRow(Number(saveBtn.getAttribute("data-spo-save") || 0)).catch((e) =>
        alert(e.message || String(e))
      );
      return;
    }
    const recvBtn = evt.target?.closest?.("button[data-spo-receive]");
    if (recvBtn) {
      receiveStoresPartOrderToInventory(Number(recvBtn.getAttribute("data-spo-receive") || 0)).catch((e) =>
        alert(e.message || String(e))
      );
      return;
    }
    const btn = evt.target?.closest?.("button[data-spo-del]");
    if (!btn) return;
    cancelStoresPartOrder(Number(btn.getAttribute("data-spo-del") || 0)).catch((e) => alert(e.message || String(e)));
  });
  qs("spoExportXlsx")?.addEventListener("click", () =>
    exportStoresPartOrdersXlsx().catch((e) => setStatus("Parts purchases Excel error: " + e.message))
  );
  qs("spoOpenPdf")?.addEventListener("click", () => openStoresPartOrdersPdf(false));
  qs("spoDownloadPdf")?.addEventListener("click", () => openStoresPartOrdersPdf(true));

  qs("ptPartsLoad")?.addEventListener("click", () =>
    loadPtPartsOrders().catch((e) => setStatus("Parts tracking error: " + e.message))
  );
  qs("ptPartsSave")?.addEventListener("click", () =>
    savePtPartsOrder().catch((e) => {
      const msg = qs("ptPartsMsg");
      if (msg) msg.textContent = e.message || String(e);
    })
  );
  qs("ptPartsClear")?.addEventListener("click", clearPtPartsForm);
  qs("ptPartsStatus")?.addEventListener("change", () => loadPtPartsOrders().catch(() => {}));
  qs("ptPartsSearch")?.addEventListener("input", () => renderPtPartsTable(ptPartsCache));
  qs("ptPartsList")?.addEventListener("change", (evt) => {
    const sel = evt.target?.closest?.("select[data-pt-status]");
    if (!sel) return;
    const id = Number(sel.getAttribute("data-pt-status") || 0);
    if (!id) return;
    const rowEl = qs("ptPartsList")?.querySelector(`[data-pt-row="${id}"]`);
    const patch = readPtPartsRowPatch(rowEl) || {};
    patch.status = String(sel.value || "").trim();
    patchStoresPartOrder(id, patch)
      .then(() => loadPtPartsOrders())
      .catch((e) => alert(e.message || String(e)));
  });
  qs("ptPartsList")?.addEventListener("click", (evt) => {
    const saveBtn = evt.target?.closest?.("button[data-pt-save]");
    if (saveBtn) {
      const id = Number(saveBtn.getAttribute("data-pt-save") || 0);
      const rowEl = qs("ptPartsList")?.querySelector(`[data-pt-row="${id}"]`);
      const patch = readPtPartsRowPatch(rowEl);
      if (!patch) return;
      patchStoresPartOrder(id, patch)
        .then(() => {
          setStatus("Part line saved.");
          return loadPtPartsOrders();
        })
        .catch((e) => alert(e.message || String(e)));
      return;
    }
    const recvBtn = evt.target?.closest?.("button[data-pt-receive]");
    if (recvBtn) {
      receiveStoresPartOrderToInventory(Number(recvBtn.getAttribute("data-pt-receive") || 0))
        .then(() => loadPtPartsOrders())
        .catch((e) => alert(e.message || String(e)));
      return;
    }
    const delBtn = evt.target?.closest?.("button[data-pt-del]");
    if (delBtn) {
      cancelStoresPartOrder(Number(delBtn.getAttribute("data-pt-del") || 0))
        .then(() => loadPtPartsOrders())
        .catch((e) => alert(e.message || String(e)));
    }
  });
  qs("ptOffLoad")?.addEventListener("click", () =>
    loadPtOffsiteRepairs().catch((e) => setStatus("Off-site tracking error: " + e.message))
  );
  qs("ptOffSave")?.addEventListener("click", () =>
    savePtOffsiteRepair().catch((e) => {
      const msg = qs("ptOffMsg");
      if (msg) msg.textContent = e.message || String(e);
    })
  );
  qs("ptOffClear")?.addEventListener("click", clearPtOffForm);
  qs("ptOffStatusFilter")?.addEventListener("change", () => renderPtOffsiteTable(ptOffsiteCache));
  qs("ptOffIncludeReturned")?.addEventListener("change", () => loadPtOffsiteRepairs().catch(() => {}));
  qs("ptOffList")?.addEventListener("click", (evt) => {
    const historyBtn = evt.target?.closest?.("button[data-pt-off-history]");
    if (historyBtn) {
      togglePtOffsiteHistory(Number(historyBtn.getAttribute("data-pt-off-history") || 0));
      return;
    }
    const btn = evt.target?.closest?.("button[data-pt-off-save]");
    if (!btn) return;
    savePtOffsiteRow(Number(btn.getAttribute("data-pt-off-save") || 0)).catch((e) =>
      alert(e.message || String(e))
    );
  });
  qs("loadAudit")?.addEventListener("click", () =>
    loadAuditLogs().catch((e) => setStatus("Audit error: " + e.message))
  );
  qs("loginSubmit")?.addEventListener("click", () => submitLoginForm());
  qs("loginPassword")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") submitLoginForm();
  });
  qs("loadAdminUsersBtn")?.addEventListener("click", () =>
    loadAdminUsers().catch((e) => setStatus("Admin users error: " + e.message))
  );
  qs("mdmLoadBtn")?.addEventListener("click", () =>
    loadMasterDataGovernance().catch((e) => setStatus("Master data error: " + e.message))
  );
  qs("mdmDeptSaveBtn")?.addEventListener("click", () =>
    saveMdmDepartment().catch((e) => setStatus("Department error: " + e.message))
  );
  qs("mdmCcSaveBtn")?.addEventListener("click", () =>
    saveMdmCostCenter().catch((e) => setStatus("Cost center error: " + e.message))
  );
  qs("mdmSupSaveBtn")?.addEventListener("click", () =>
    saveMdmSupplier().catch((e) => setStatus("Supplier error: " + e.message))
  );
  qs("mdmPolicySaveBtn")?.addEventListener("click", () =>
    saveMdmPolicies().catch((e) => setStatus("Policy error: " + e.message))
  );
  qs("adminArtisanPresetBtn")?.addEventListener("click", applyAdminArtisanPreset);
  qs("saveAdminUserBtn")?.addEventListener("click", () =>
    saveAdminUser().catch((e) => setStatus("Save user error: " + e.message))
  );
  qs("chPwdSubmit")?.addEventListener("click", () =>
    submitChangePassword().catch((e) => setStatus("Password error: " + e.message))
  );
  qs("loadSmtpSettingsBtn")?.addEventListener("click", () =>
    loadSmtpSettings().catch((e) => setStatus("SMTP load error: " + e.message))
  );
  qs("saveSmtpSettingsBtn")?.addEventListener("click", () =>
    saveSmtpSettings().catch((e) => setStatus("SMTP save error: " + e.message))
  );
  qs("testSmtpSettingsBtn")?.addEventListener("click", () =>
    testSmtpSettings().catch((e) => setStatus("SMTP test error: " + e.message))
  );
  qs("loadPushNotifyBtn")?.addEventListener("click", () =>
    loadPushNotificationSettings().catch((e) => setStatus("Push load error: " + e.message))
  );
  qs("sendPushNotifyTestBtn")?.addEventListener("click", () =>
    sendPushNotificationTest().catch((e) => setStatus("Push test error: " + e.message))
  );
  qs("sendPushNotifyBtn")?.addEventListener("click", () =>
    sendPushNotificationManual().catch((e) => setStatus("Push send error: " + e.message))
  );
  qs("loadPdfReportSettingsBtn")?.addEventListener("click", () =>
    loadPdfReportSettings().catch((e) => setStatus("PDF site load error: " + e.message))
  );
  qs("savePdfReportSettingsBtn")?.addEventListener("click", () =>
    savePdfReportSettings().catch((e) => setStatus("PDF site save error: " + e.message))
  );
  qs("uploadPdfReportLogoBtn")?.addEventListener("click", () =>
    uploadPdfReportLogo().catch((e) => setStatus("PDF logo upload error: " + e.message))
  );
  qs("removePdfReportLogoBtn")?.addEventListener("click", () =>
    removePdfReportLogo().catch((e) => setStatus("PDF logo remove error: " + e.message))
  );
  qs("pdfReportSiteCode")?.addEventListener("change", () => onPdfReportSiteCodeChange());
  qs("pdfReportCompanyCode")?.addEventListener("change", () => onPdfReportCompanyCodeChange());
  qs("sendSubscriptionNowBtn")?.addEventListener("click", () =>
    sendSubscriptionNowFromAdmin().catch((e) => setStatus("Subscription send error: " + e.message))
  );
  qs("refreshBackupsBtn")?.addEventListener("click", () =>
    loadBackupFiles().catch((e) => setStatus("Backups load error: " + e.message))
  );
  qs("createBackupNowBtn")?.addEventListener("click", () =>
    createBackupNow().catch((e) => setStatus("Backup create error: " + e.message))
  );
  qs("previewBackupRestoreBtn")?.addEventListener("click", () =>
    previewBackupRestore().catch((e) => setStatus("Backup preview error: " + e.message))
  );
  qs("stageBackupRestoreBtn")?.addEventListener("click", () =>
    stageBackupRestore().catch((e) => setStatus("Restore stage error: " + e.message))
  );
  qs("executeBackupRestoreBtn")?.addEventListener("click", () =>
    executeBackupRestoreNow().catch((e) => setStatus("Restore execute error: " + e.message))
  );
  qs("loadApprovals")?.addEventListener("click", () =>
    loadApprovalRequests().catch((e) => setStatus("Approvals error: " + e.message))
  );
  qs("approvalStatus")?.addEventListener("change", () =>
    loadApprovalRequests().catch((e) => setStatus("Approvals error: " + e.message))
  );
  qs("legalUploadBtn")?.addEventListener("click", () =>
    uploadLegalDoc().catch((e) => setStatus("Legal upload error: " + e.message))
  );
  qs("loadLegalBtn")?.addEventListener("click", () =>
    loadLegalDocs().catch((e) => setStatus("Legal load error: " + e.message))
  );
  qs("loadLegalExpiryBtn")?.addEventListener("click", () =>
    loadLegalExpiry().catch((e) => setStatus("Legal expiry error: " + e.message))
  );
  qs("openLegalCompliancePdf")?.addEventListener("click", () => openLegalCompliancePdf(false));
  qs("downloadLegalCompliancePdf")?.addEventListener("click", () => openLegalCompliancePdf(true));
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

  qs("openDaily")?.addEventListener("click", openDailyPdf);
  qs("openWeekly")?.addEventListener("click", openWeeklyPdf);
  qs("downloadAmlWeeklyCheckSheet")?.addEventListener("click", () => {
    downloadAmlWeeklyCheckSheet().catch((e) => {
      const status = qs("amlWeeklyExportStatus");
      if (status) status.textContent = `Export failed: ${e.message || e}`;
      setStatus("AML Weekly Check Sheet export error: " + (e.message || e));
    });
  });
  qs("openLubePdf")?.addEventListener("click", openLubePdf);
  qs("openLubePdfFromLube")?.addEventListener("click", openLubePdf);
  qs("downloadLubeUsageXlsx")?.addEventListener("click", downloadLubeUsageXlsx);
  qs("loadLubeMonthStock")?.addEventListener("click", () =>
    loadLubeMonthStock().catch((e) => setStatus("Lube month stock error: " + e.message))
  );
  qs("openStockMonitorPdf")?.addEventListener("click", openStockMonitorPdf);
  qs("downloadStockMonitorPdf")?.addEventListener("click", downloadStockMonitorPdf);
  qs("openOperationsPdf")?.addEventListener("click", () => openOperationsPdf(false));
  qs("downloadOperationsPdf")?.addEventListener("click", () => openOperationsPdf(true));
  qs("downloadOperationsXlsx")?.addEventListener("click", downloadOperationsXlsx);
  qs("openDailyXlsx")?.addEventListener("click", openDailyXlsx);
  qs("openGmWeeklyXlsx")?.addEventListener("click", openGmWeeklyXlsx);
  qs("downloadCostMonthlyXlsx")?.addEventListener("click", downloadCostMonthlyXlsx);
  qs("openMonthlyFleetCostPdf")?.addEventListener("click", openMonthlyFleetCostPdf);
  qs("downloadMonthlyFleetCostPdf")?.addEventListener("click", downloadMonthlyFleetCostPdf);
  qs("downloadMtdOpeningHoursXlsx")?.addEventListener("click", downloadMtdOpeningHoursXlsx);
  qs("downloadMaintenanceCostByEquipmentXlsx")?.addEventListener("click", downloadMaintenanceCostByEquipmentXlsx);
  qs("openMaintenanceCostByEquipmentPdf")?.addEventListener("click", () => openMaintenanceCostByEquipmentPdf(false));
  qs("downloadMaintenanceCostByEquipmentPdf")?.addEventListener("click", () => openMaintenanceCostByEquipmentPdf(true));
  qs("downloadMaintenanceExecutivePptx")?.addEventListener("click", downloadMaintenanceExecutivePptx);
  qs("downloadGMUpcomingCostsPptx")?.addEventListener("click", downloadGMUpcomingCostsPptx);
  qs("downloadGMBudgetMeetingDocx")?.addEventListener("click", downloadGMBudgetMeetingDocx);
  qs("saveRainDayBtn")?.addEventListener("click", () => saveRainDay().catch((e) => setStatus("Rain day save error: " + e.message)));
  qs("removeRainDayBtn")?.addEventListener("click", () => removeRainDay().catch((e) => setStatus("Rain day remove error: " + e.message)));
  qs("loadRainDaysBtn")?.addEventListener("click", () => loadRainDays().catch((e) => setStatus("Rain day load error: " + e.message)));

  qs("boRefreshOpen")?.addEventListener("click", () =>
    loadBreakdownOpsOpen().catch((e) => setStatus("Open list error: " + e.message))
  );
  qs("boRefreshRecent")?.addEventListener("click", () =>
    loadBreakdownOpsRecent().catch((e) => setStatus("Recent list error: " + e.message))
  );
  qs("boEnsureOpen")?.addEventListener("click", () =>
    ensureOpenBreakdownOps().catch((e) => setStatus("Ensure open error: " + e.message))
  );
  qs("boRepairCreateWo")?.addEventListener("click", () =>
    createRepairWorkOrderOps().catch((e) => setStatus("Repair WO error: " + e.message))
  );
  qs("boRepairOpenWo")?.addEventListener("click", () => {
    if (lastBoRepairWoId) window.open(`/web/workorders.html?wo=${encodeURIComponent(String(lastBoRepairWoId))}`, "_blank");
  });
  qs("boPullLiveHours")?.addEventListener("click", () =>
    pullBreakdownOpsLiveHours().catch((e) => setStatus("Live hours error: " + e.message))
  );
  qs("boOpenList")?.addEventListener("click", (ev) => {
    const w = ev.target?.closest?.(".bo-copy-wo");
    if (w) {
      const wo = w.getAttribute("data-wo");
      if (wo && qs("iWo")) qs("iWo").value = String(wo);
      setStatus(`Copied WO #${wo} for parts issue.`);
      return;
    }
    const c = ev.target?.closest?.(".bo-close-bdn");
    if (c) closeBreakdownFromOps(c.getAttribute("data-id")).catch(() => {});
  });

  qs("boSlipType")?.addEventListener("change", updateBoSlipFormVisibility);
  qs("boSlipPhotosInput")?.addEventListener("change", (e) =>
    onBoSlipPhotosInputChange(e).catch((err) => setStatus(String(err.message || err)))
  );
  qs("boSlipPhotosClear")?.addEventListener("click", () => {
    clearBoSlipPhotosUi();
    setStatus("Slip pictures cleared.");
  });
  qs("boSlipPullAsset")?.addEventListener("click", () =>
    pullBoSlipFromAsset().catch((e) => setStatus("Pull from asset error: " + (e.message || e)))
  );
  qs("boSlipSave")?.addEventListener("click", () =>
    saveBoSlipReport().catch((e) => setStatus("Slip save error: " + e.message))
  );
  qs("boSlipLoadList")?.addEventListener("click", () =>
    loadBoSlipSavedList().catch((e) => setStatus("Slip list error: " + e.message))
  );
  qs("boSlipSavedList")?.addEventListener("click", (ev) => {
    const b = ev.target?.closest?.(".bo-slip-pdf");
    if (b) openBoSlipPdf(b.getAttribute("data-id"));
  });

  qs("makeBreakdown")?.addEventListener("click", () =>
    createBreakdown().catch((e) => setStatus("Breakdown error: " + e.message))
  );
  bindShortBreakdownPartsUi();
  qs("sqSubmit")?.addEventListener("click", () =>
    submitShortBreakdown().catch((e) => setStatus("Short breakdown error: " + e.message))
  );
  qs("issuePart")?.addEventListener("click", () =>
    issuePart().catch((e) => setStatus("Issue error: " + e.message))
  );
  qs("allocateStore")?.addEventListener("click", () =>
    allocateStore().catch((e) => setStatus("Stores allocation error: " + e.message))
  );
  qs("refreshAllocations")?.addEventListener("click", () =>
    loadStoreAllocations().catch((e) => setStatus("Allocation list error: " + e.message))
  );
  qs("saveManualStock")?.addEventListener("click", () =>
    saveManualStock().catch((e) => setStatus("Manual stock error: " + e.message))
  );
  qs("msPart")?.addEventListener("input", updateManualStockPartDesc);
  qs("msPart")?.addEventListener("change", updateManualStockPartDesc);
  qs("msType")?.addEventListener("change", () => {
    updateManualStockCostRowVisibility();
  });
  qs("msType")?.addEventListener("input", () => {
    updateManualStockCostRowVisibility();
  });
  qs("mlPart")?.addEventListener("input", updateManualLubePartDesc);
  qs("mlPart")?.addEventListener("change", updateManualLubePartDesc);
  // Lube minimums moved to separate card
  qs("lubeMinPart")?.addEventListener("input", updateLubeMinPartDesc);
  qs("lubeMinPart")?.addEventListener("change", updateLubeMinPartDesc);
  qs("lubeMinSetOne")?.addEventListener("click", () =>
    setSingleLubeMinimum().catch((e) => setStatus("Lube min error: " + e.message))
  );
  qs("lubeMinRefresh")?.addEventListener("click", () =>
    loadLubeReorderAlerts().catch((e) => setStatus("Lube alerts error: " + e.message))
  );
  qs("receiveLube")?.addEventListener("click", () =>
    receiveLubeStock().catch((e) => setStatus("Receive lube error: " + e.message))
  );
  qs("lrPart")?.addEventListener("input", updateReceiveLubePartDesc);
  qs("lrPart")?.addEventListener("change", updateReceiveLubePartDesc);
  qs("icLoad")?.addEventListener("click", () =>
    loadInventoryControl().catch((e) => setStatus("Inventory control error: " + e.message))
  );
  qs("icSaveMin")?.addEventListener("click", () =>
    saveInventoryPartMinimum().catch((e) => setStatus("Part minimum error: " + e.message))
  );
  qs("icSubmitCount")?.addEventListener("click", () =>
    submitInventoryCycleCount().catch((e) => setStatus("Cycle count error: " + e.message))
  );
  qs("icPartCode")?.addEventListener("change", () =>
    loadInventoryControl().catch((e) => setStatus("Inventory control error: " + e.message))
  );
  qs("saveManualLube")?.addEventListener("click", () =>
    saveManualLube().catch((e) => setStatus("Manual lube error: " + e.message))
  );
  ["msLocation", "saLocation", "mlLocation"].forEach((id) => {
    qs(id)?.addEventListener("change", () => {
      const v = String(qs(id)?.value || "").trim().toUpperCase();
      if (!v) return;
      setRoleDefaultLocation(getSessionRole(), v);
      applyDefaultLocationsToInputs();
      if (id === "saLocation") {
        const binInput = qs("saBin");
        if (binInput) binInput.value = "";
        loadBinCodeOptionsForLocation(v, "saBinCodeOptions").catch(() => {});
      }
      if (id === "msLocation") {
        const binInput = qs("msBin");
        if (binInput) binInput.value = "";
        loadBinCodeOptionsForLocation(v, "msBinCodeOptions").catch(() => {});
      }
    });
  });
  qs("locLoad")?.addEventListener("click", () =>
    loadLocations().catch((e) => setStatus("Locations error: " + e.message))
  );
  qs("locShowInactive")?.addEventListener("change", () =>
    loadLocations().catch((e) => setStatus("Locations error: " + e.message))
  );
  qs("locSave")?.addEventListener("click", () =>
    saveLocation().catch((e) => setStatus("Location save error: " + e.message))
  );
  qs("sbSaveBinBtn")?.addEventListener("click", () =>
    saveStockBin().catch((e) => setStatus("Bin save error: " + e.message))
  );
  qs("sbLoadBinsBtn")?.addEventListener("click", () =>
    loadStockBins().catch((e) => setStatus("Bins load error: " + e.message))
  );
  qs("sbSaveMinMaxBtn")?.addEventListener("click", () =>
    saveStockMinMax().catch((e) => setStatus("Min-max save error: " + e.message))
  );
  qs("sbLoadDepthBtn")?.addEventListener("click", () =>
    loadStockDepth().catch((e) => setStatus("Depth load error: " + e.message))
  );
  qs("sbLoadReplenishmentBtn")?.addEventListener("click", () =>
    loadReplenishmentSuggestions().catch((e) => setStatus("Replenishment load error: " + e.message))
  );
  qs("sbExportReplenishmentCsvBtn")?.addEventListener("click", () => {
    try {
      exportReplenishmentSuggestionsCsv();
      setStatus("Replenishment CSV exported.");
    } catch (e) {
      setStatus("Replenishment export error: " + (e.message || e));
    }
  });
  qs("sbCreateCycleSessionBtn")?.addEventListener("click", () =>
    createCycleSession().catch((e) => setStatus("Cycle session create error: " + e.message))
  );
  qs("sbLoadCycleSessionsBtn")?.addEventListener("click", () =>
    loadCycleSessions().catch((e) => setStatus("Cycle sessions load error: " + e.message))
  );
  qs("loadLubeStock")?.addEventListener("click", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlPart")?.addEventListener("change", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlPart")?.addEventListener("input", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlType")?.addEventListener("change", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlType")?.addEventListener("input", () =>
    loadLubeStockOnHand().catch((e) => setStatus("Lube stock error: " + e.message))
  );
  qs("mlQty")?.addEventListener("input", updateLubeQtyWarning);
  qs("setLubeMin210")?.addEventListener("click", () =>
    setLubeMinimumStock().catch((e) => setStatus("Lube minimum error: " + e.message))
  );

  // Daily
  qs("loadDaily")?.addEventListener("click", () =>
    loadDailyInput().catch((e) => setStatus("Daily load error: " + e.message))
  );
  qs("saveDaily")?.addEventListener("click", () =>
    saveDailyInput().catch((e) => setStatus("Daily save error: " + e.message))
  );
  qs("scheduled")?.addEventListener("input", () => {
    const raw = toNum(qs("scheduled")?.value);
    if (raw == null || raw <= 0 || raw > 24) return;
    applyDayScheduledHours(raw);
    setStatus(`Production hours updated to ${raw} in Daily Log.`);
  });
  qs("scheduled")?.addEventListener("change", () => {
    const scheduled = getDayScheduledHours();
    if (qs("scheduled")) qs("scheduled").value = String(scheduled);
    applyDayScheduledHours(scheduled);
    loadDashboard()
      .then(() => setStatus(`Production hours set to ${scheduled} for the selected day.`))
      .catch((e) => setStatus("Scheduled hours update error: " + e.message));
  });
  qs("runShiftSelfCheck")?.addEventListener("click", () =>
    runShiftSelfCheck().catch((e) => setStatus("Self-check error: " + e.message))
  );
  qs("exportShiftSelfCheck")?.addEventListener("click", exportShiftSelfCheckTxt);

  qs("copyYesterday")?.addEventListener("click", () =>
    copyYesterdayToToday().catch((e) => setStatus("Copy yesterday error: " + e.message))
  );
  qs("dailyHoursCsvBtn")?.addEventListener("click", () => qs("dailyHoursCsvFile")?.click());
  qs("dailyHoursCsvTemplate")?.addEventListener("click", downloadDailyHoursCsvTemplate);
  qs("dailyHoursCsvFile")?.addEventListener("change", (e) => {
    const file = e.target?.files?.[0];
    if (file) uploadDailyHoursCsv(file).finally(() => { e.target.value = ""; });
  });
  qs("dailyMatrixCsvBtn")?.addEventListener("click", () => qs("dailyMatrixCsvFile")?.click());
  qs("dailyMatrixCsvTemplate")?.addEventListener("click", downloadDailyMatrixCsvTemplate);
  qs("dailyMatrixCsvFile")?.addEventListener("change", (e) => {
    const file = e.target?.files?.[0];
    if (file) uploadDailyMatrixCsv(file).finally(() => { e.target.value = ""; });
  });
  qs("applyBulkSched")?.addEventListener("click", applyBulkScheduled);
  qs("dailyDownOnly")?.addEventListener("change", () => {
    dailyShowDownOnly = !!qs("dailyDownOnly")?.checked;
    renderDailyTable();
  });
  qs("dailyQrGenerate")?.addEventListener("click", () =>
    generateDailyAssetQr().catch((e) => setStatus("QR generate error: " + e.message))
  );
  qs("dailyQrPrint")?.addEventListener("click", printDailyAssetQr);
  qs("dailyQrDownloadVisible")?.addEventListener("click", () =>
    downloadAllVisibleDailyQrs().catch((e) => setStatus("Bulk QR download error: " + e.message))
  );
  qs("dailyQrPrintVisible")?.addEventListener("click", () =>
    printVisibleDailyQrSheet().catch((e) => setStatus("QR sheet print error: " + e.message))
  );
  qs("qrPreset")?.addEventListener("change", applyQrSheetPreset);
  ["qrCols", "qrSizeMm", "qrCellMm", "qrGapMm"].forEach((id) => {
    qs(id)?.addEventListener("input", () => {
      const preset = qs("qrPreset");
      if (preset && preset.value !== "custom") preset.value = "custom";
    });
  });
  applyQrSheetPreset();

  qs("safetyTplLoadBtn")?.addEventListener("click", () =>
    loadSafetyTemplateEditor().catch((e) => setStatus("Safety template error: " + e.message))
  );
  qs("safetyTplSelect")?.addEventListener("change", () =>
    loadSafetyTemplateEditor().catch(() => {})
  );
  qs("safetyCategoryAddBtn")?.addEventListener("click", () =>
    addSafetyCategory().catch((e) => setStatus("Add category error: " + e.message))
  );
  qs("safetyCategoriesList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-safety-edit-category]");
    if (!btn) return;
    const key = String(btn.getAttribute("data-safety-edit-category") || "").trim();
    if (qs("safetyTplSelect")) qs("safetyTplSelect").value = key;
    loadSafetyTemplateEditor().catch(() => {});
  });
  qs("safetyTplSaveBtn")?.addEventListener("click", () =>
    saveSafetyTemplateEditor().catch((e) => setStatus("Safety template save error: " + e.message))
  );
  qs("safetyTplAddRowBtn")?.addEventListener("click", () => {
    safetyTplItems.push({ key: `item_${safetyTplItems.length + 1}`, label: "New checklist item" });
    renderSafetyTemplateEditor(safetyTplItems);
  });
  qs("safetyTplItems")?.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-safety-tpl-remove]");
    if (!btn) return;
    const idx = Number(btn.getAttribute("data-safety-tpl-remove"));
    safetyTplItems.splice(idx, 1);
    renderSafetyTemplateEditor(safetyTplItems);
  });
  qs("safetyItemAddBtn")?.addEventListener("click", () =>
    addSafetyEquipmentItem().catch((e) => setStatus("Add safety item error: " + e.message))
  );
  qs("safetyItemsList")?.addEventListener("click", (e) => {
    const qrBtn = e.target.closest("button[data-safety-use-qr]");
    if (qrBtn) {
      const code = String(qrBtn.getAttribute("data-safety-use-qr") || "");
      if (qs("safetyQrItemCode")) qs("safetyQrItemCode").value = code;
      generateSafetyQr().catch((err) => setStatus("QR error: " + err.message));
      return;
    }
    const inspBtn = e.target.closest("button[data-safety-open-insp]");
    if (inspBtn) {
      const code = String(inspBtn.getAttribute("data-safety-open-insp") || "");
      if (code) window.open(`./safety-inspection.html?item_code=${encodeURIComponent(code)}`, "_blank");
      return;
    }
    const pdfBtn = e.target.closest("button[data-safety-item-pdf]");
    if (pdfBtn) {
      openSafetyItemInspectionPdf(pdfBtn.getAttribute("data-safety-item-pdf"))
        .catch((err) => setStatus("Safety PDF error: " + err.message));
      return;
    }
    const rmBtn = e.target.closest("button[data-safety-remove-item]");
    if (rmBtn) {
      removeSafetyEquipmentItem(Number(rmBtn.getAttribute("data-safety-remove-item") || 0))
        .catch((err) => setStatus("Remove error: " + err.message));
    }
  });
  qs("safetyItemsList")?.addEventListener("change", (e) => {
    const chk = e.target.closest("input[data-safety-report-select]");
    if (!chk) return;
    const code = String(chk.getAttribute("data-safety-report-select") || "").trim().toUpperCase();
    if (!code) return;
    if (chk.checked) safetyReportSelectedCodes.add(code);
    else safetyReportSelectedCodes.delete(code);
  });
  qs("safetyQrGenerate")?.addEventListener("click", () =>
    generateSafetyQr().catch((e) => setStatus("Safety QR error: " + e.message))
  );
  qs("safetyQrPrint")?.addEventListener("click", printSafetyQr);
  qs("safetyQrPrintSheet")?.addEventListener("click", () =>
    printAllSafetyQrSheet().catch((e) => setStatus("Safety QR sheet error: " + e.message))
  );
  qs("safetyQrPreset")?.addEventListener("change", applySafetyQrSheetPreset);
  ["safetyQrCols", "safetyQrSizeMm", "safetyQrCellMm", "safetyQrGapMm"].forEach((id) => {
    qs(id)?.addEventListener("input", () => {
      const preset = qs("safetyQrPreset");
      if (preset && preset.value !== "custom") preset.value = "custom";
    });
  });
  qs("safetyPdfRegisterBtn")?.addEventListener("click", () =>
    openSafetyRegisterPdf(false).catch((e) => setStatus("Safety PDF error: " + e.message))
  );
  qs("safetyPdfBlankBtn")?.addEventListener("click", () =>
    openSafetyRegisterPdf(true).catch((e) => setStatus("Safety PDF error: " + e.message))
  );
  qs("safetyInspectionReportAllBtn")?.addEventListener("click", () =>
    openSafetyInspectionReportPdf(false).catch((e) => setStatus("Safety report error: " + e.message))
  );
  qs("safetyInspectionReportSelectedBtn")?.addEventListener("click", () =>
    openSafetyInspectionReportPdf(true).catch((e) => setStatus("Safety report error: " + e.message))
  );
  initSafetyAdminPanel().catch(() => {});
  initOfflineQueueAdminPanel();
  initTelematicsAdminPanel().catch(() => {});
  initCartrackAdminPanel().catch(() => {});
  initCartrackTrackingTab();

  qs("cartrackSaveSettingsBtn")?.addEventListener("click", () =>
    saveCartrackAdminSettings().catch((e) => setCartrackAdminResult(String(e.message || e), false))
  );
  qs("cartrackTestBtn")?.addEventListener("click", () => testCartrackConnection().catch(() => {}));
  qs("cartrackRunMorningBtn")?.addEventListener("click", () => runCartrackMorningNow().catch(() => {}));
  qs("unitechSaveSettingsBtn")?.addEventListener("click", () =>
    saveUnitechAdminSettings().catch((e) => setUnitechAdminResult(String(e.message || e), false))
  );
  qs("unitechTestBtn")?.addEventListener("click", () => testUnitechConnection().catch(() => {}));
  qs("gpsLinkSaveBtn")?.addEventListener("click", () => saveGpsVehicleLink().catch(() => {}));
  qs("gpsLinkRefreshBtn")?.addEventListener("click", () => loadGpsVehicleLinksAdmin().catch(() => {}));
  qs("gpsLinkApplyBtn")?.addEventListener("click", () => applyGpsVehicleLinks().catch(() => {}));
  qs("gpsVehicleLinksList")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-gps-link-delete]");
    if (!btn) return;
    deleteGpsVehicleLink(btn.getAttribute("data-gps-link-delete")).catch(() => {});
  });
  qs("gpsVehicleLinkSuggestions")?.addEventListener("click", (e) => {
    const btn = e.target.closest("[data-gps-link-prefill]");
    if (!btn) return;
    prefillGpsVehicleLinkForm(btn.getAttribute("data-gps-link-prefill"), btn.getAttribute("data-gps-link-source"));
  });
  qs("cartrackSyncBtn")?.addEventListener("click", () => syncCartrackNow().catch((e) => setStatus(String(e.message || e))));
  qs("cartrackDashShowAll")?.addEventListener("change", () => {
    renderCartrackFleetTable(
      cartrackDashboardFleetCache.fleet,
      cartrackDashboardFleetCache.speedingToday,
      { showAll: Boolean(qs("cartrackDashShowAll")?.checked) }
    );
  });
  qs("cartrackMorningPdfBtn")?.addEventListener("click", () =>
    openCartrackMorningPdf().catch((e) => setStatus(String(e.message || e)))
  );
  qs("cartrackMorningEmailBtn")?.addEventListener("click", () =>
    emailCartrackMorningReport().catch((e) => setStatus(String(e.message || e)))
  );
  qs("cartrackTrackMorningPdfBtn")?.addEventListener("click", () =>
    openCartrackMorningPdf().catch((e) => setStatus(String(e.message || e)))
  );
  const onCartrackSpeedReportDateChange = (e) => {
    const d = syncCartrackSpeedReportDateInputs(e.target?.value);
    loadCartrackSpeedingEvents(d, { useCache: d === todayLocalYmd() }).catch(() => {});
    loadCartrackSpeedingEvents(d, {
      hostId: "cartrackTrackSpeedingEventsHost",
      countId: "cartrackTrackSpeedingEventsCount",
      panelId: "cartrackTrackSpeedingEventsPanel",
      useCache: d === todayLocalYmd(),
    }).catch(() => {});
  };
  qs("cartrackSpeedReportDate")?.addEventListener("change", onCartrackSpeedReportDateChange);
  qs("cartrackTrackSpeedReportDate")?.addEventListener("change", onCartrackSpeedReportDateChange);
  qs("cartrackKpiSpeeding")?.closest(".kpi-pill")?.addEventListener("click", () => {
    openCartrackSpeedingEventsPanel();
    loadCartrackSpeedingEvents(todayLocalYmd(), { useCache: true }).catch(() => {});
  });
  initCartrackSpeedReportDates();

  qs("telemSaveDeviceBtn")?.addEventListener("click", () =>
    saveTelematicsDevice().catch((e) => setTelemAdminResult(String(e.message || e)))
  );
  qs("telemRefreshDevicesBtn")?.addEventListener("click", () =>
    loadTelematicsAdminDevices().catch((e) => setTelemAdminResult(String(e.message || e)))
  );
  qs("telemShowInactive")?.addEventListener("change", () =>
    loadTelematicsAdminDevices().catch(() => {})
  );
  qs("telematicsRefreshBtn")?.addEventListener("click", () =>
    loadTelematicsTab().catch((e) => setStatus("Telematics refresh error: " + (e.message || e)))
  );
  qs("telemDevicesList")?.addEventListener("click", (e) => {
    const editBtn = e.target.closest("button[data-telem-edit]");
    if (editBtn) {
      fillTelematicsDeviceForm({
        assetCode: editBtn.getAttribute("data-telem-asset"),
        deviceSerial: "",
        unitModel: editBtn.getAttribute("data-telem-model") || "FSC650",
        externalId: "",
        replaceFaulty: true,
      });
      setTelemAdminResult(`Enter new serial for ${editBtn.getAttribute("data-telem-asset")} (was ${editBtn.getAttribute("data-telem-serial")}).`);
      return;
    }
    const deactBtn = e.target.closest("button[data-telem-deactivate]");
    if (deactBtn) {
      deactivateTelematicsDeviceAdmin(
        deactBtn.getAttribute("data-telem-deactivate"),
        deactBtn.getAttribute("data-telem-asset-label")
      ).catch((err) => setTelemAdminResult(String(err.message || err)));
    }
  });

  // Net banner
  refreshNetBanner();
  window.addEventListener("offline", () => {
    refreshNetBanner();
    renderOfflineQueueAdminPanel();
  });

  window.addEventListener("online", async () => {
    refreshNetBanner();
    renderOfflineQueueAdminPanel();
    if (getTotalQueuedCount() === 0) return;
    try {
      await syncAllOfflineQueues();
    } catch (e) {
      setStatus("Sync error: " + (e.message || e));
      refreshNetBanner();
    }
  });

  qs("syncNow")?.addEventListener("click", async () => {
    if (!navigator.onLine) return alert("Still offline.");
    try {
      await syncAllOfflineQueues();
    } catch (e) {
      setStatus("Sync error: " + (e.message || e));
      refreshNetBanner();
    }
  });

  // Assets
  qs("loadHistory")?.addEventListener("click", () =>
    loadAssetHistory().catch((e) => setStatus("History error: " + e.message))
  );
  qs("downloadHistoryPdf")?.addEventListener("click", downloadAssetHistoryPdf);

  qs("historyList")?.addEventListener("click", async (e) => {
    const btn = e.target.closest("button[data-ops-slip-pdf]");
    if (!btn) return;
    e.preventDefault();
    const id = btn.getAttribute("data-ops-slip-pdf");
    if (!id) return;
    try {
      setStatus("Opening slip PDF...");
      await openAuthedPdf(`${API}/api/breakdown-ops/slips/${encodeURIComponent(id)}/pdf`);
      setStatus("PDF opened ✅");
    } catch (err) {
      setStatus("PDF error: " + (err.message || err));
    }
  });

  qs("showArchived")?.addEventListener("change", () => {
    loadAssetsFleet().catch(() => {});
  });

  qs("downloadAssetsCostCentersXlsx")?.addEventListener("click", () => downloadAssetsCostCentersXlsx());

  qs("saveOperatingBudget")?.addEventListener("click", () =>
    saveOperatingBudget().catch((e) => setStatus("Operating budget error: " + e.message))
  );
  qs("savePlantHireBudget")?.addEventListener("click", () =>
    savePlantHireBudget().catch((e) => setStatus("Plant hire budget error: " + e.message))
  );
  qs("savePlantHireRates")?.addEventListener("click", () =>
    savePlantHireRates().catch((e) => setStatus("Plant hire rates error: " + e.message))
  );
  qs("plantHireBudgetMonth")?.addEventListener("change", () => loadPlantHireBudgetStatus().catch(() => {}));
  qs("plantHireAssetSelect")?.addEventListener("change", () => {
    const code = String(qs("plantHireAssetSelect")?.value || "").trim();
    const row = plantHireRegisterCache.find((r) => r.asset_code === code);
    fillPlantHireRateFields(row || {});
    syncPlantHireAssetLabel(code);
  });

  qs("assetsFleetFilter")?.addEventListener("input", () => {
    renderAssetFleetGrid(assetsFleetCache);
  });
  qs("assetsFleetStatus")?.addEventListener("change", () => renderAssetFleetGrid(assetsFleetCache));
  qs("assetFleetCard")?.addEventListener("click", (e) => {
    const filterBtn = e.target.closest("button[data-fleet-status]");
    if (!filterBtn) return;
    if (qs("assetsFleetStatus")) qs("assetsFleetStatus").value = filterBtn.dataset.fleetStatus || "";
    renderAssetFleetGrid(assetsFleetCache);
  });

  qs("assetFleetGrid")?.addEventListener("click", (e) => {
    const card = e.target.closest(".asset-fleet-card");
    if (!card) return;
    const code = card.dataset.assetCode;
    if (!code) return;
    selectAssetCard(code, { loadHistory: true, scroll: true }).catch((err) =>
      setStatus("Asset select error: " + (err.message || err))
    );
  });

  qs("btnArchiveAsset")?.addEventListener("click", () =>
    archiveSelectedAsset().catch((e) => setStatus("Archive error: " + e.message))
  );

  qs("btnUnarchiveAsset")?.addEventListener("click", () =>
    unarchiveSelectedAsset().catch((e) => setStatus("Unarchive error: " + e.message))
  );
  qs("saveAssetDetails")?.addEventListener("click", () =>
    saveAssetDetails().catch((e) => setStatus("Equipment details error: " + e.message))
  );
  qs("saveContractorAsset")?.addEventListener("click", () =>
    saveContractorAsset().catch((e) => setStatus("Contractor asset save error: " + e.message))
  );
  qs("assetAllocSelect")?.addEventListener("change", () => loadAssetAllocationForm());
  qs("saveAssetAllocationBtn")?.addEventListener("click", () =>
    saveAssetAllocation().catch((e) => setStatus("Asset allocation error: " + e.message))
  );
  qs("refreshAssetAllocationBtn")?.addEventListener("click", () =>
    populateAssetAllocSelect().catch((e) => setStatus("Asset allocation refresh error: " + e.message))
  );
  populateAssetAllocSelect().catch(() => {});

  loadAssetsFleet().catch(() => {});
  const saDate = qs("saDate");
  if (saDate) saDate.value = new Date().toISOString().slice(0, 10);
  const mlDate = qs("mlDate");
  if (mlDate) mlDate.value = new Date().toISOString().slice(0, 10);
  const fuelDate = qs("fuelDate");
  if (fuelDate) fuelDate.value = new Date().toISOString().slice(0, 10);
  const lubeRange = getDefaultLubeRange();
  const lubeStart = qs("lubeStart");
  const lubeEnd = qs("lubeEnd");
  if (lubeStart && !lubeStart.value) lubeStart.value = lubeRange.start;
  if (lubeEnd && !lubeEnd.value) lubeEnd.value = lubeRange.end;
  const lubeStockMonth = qs("lubeStockMonth");
  if (lubeStockMonth && !lubeStockMonth.value) {
    lubeStockMonth.value = (lubeStart?.value || lubeRange.end).slice(0, 7);
  }
  const fuelStart = qs("fuelStart");
  const fuelEnd = qs("fuelEnd");
  if (fuelStart && !fuelStart.value) fuelStart.value = lubeRange.start;
  if (fuelEnd && !fuelEnd.value) fuelEnd.value = lubeRange.end;
  const fuelTolerance = qs("fuelTolerance");
  if (fuelTolerance && !fuelTolerance.value) fuelTolerance.value = "0.15";
  const fuelSnapStart = qs("fuelSnapStart");
  const fuelSnapEnd = qs("fuelSnapEnd");
  if (fuelSnapStart && !fuelSnapStart.value) fuelSnapStart.value = fuelStart?.value || lubeRange.start;
  if (fuelSnapEnd && !fuelSnapEnd.value) fuelSnapEnd.value = fuelEnd?.value || date.value;
  const opDate = qs("opDate");
  if (opDate && !opDate.value) opDate.value = new Date().toISOString().slice(0, 10);
  const opSiteDate = qs("opSiteDate");
  if (opSiteDate && !opSiteDate.value) opSiteDate.value = new Date().toISOString().slice(0, 10);
  const opSiteDashDate = qs("opSiteDashDate");
  if (opSiteDashDate && !opSiteDashDate.value) opSiteDashDate.value = opSiteDate?.value || new Date().toISOString().slice(0, 10);
  const opDelayDate = qs("opDelayDate");
  if (opDelayDate && !opDelayDate.value) opDelayDate.value = new Date().toISOString().slice(0, 10);
  const opTargetDate = qs("opTargetDate");
  if (opTargetDate && !opTargetDate.value) opTargetDate.value = new Date().toISOString().slice(0, 10);
  const opFrom = qs("opFrom");
  const opTo = qs("opTo");
  if (opFrom && !opFrom.value) opFrom.value = new Date(Date.now() - 1000 * 60 * 60 * 24 * 30).toISOString().slice(0, 10);
  if (opTo && !opTo.value) opTo.value = new Date().toISOString().slice(0, 10);
  const dpDate = qs("dpDate");
  if (dpDate && !dpDate.value) dpDate.value = new Date().toISOString().slice(0, 10);
  const dpFrom = qs("dpFrom");
  const dpTo = qs("dpTo");
  if (dpFrom && !dpFrom.value) dpFrom.value = new Date(Date.now() - 1000 * 60 * 60 * 24 * 7).toISOString().slice(0, 10);
  if (dpTo && !dpTo.value) dpTo.value = new Date().toISOString().slice(0, 10);
  const qFrom = qs("qFrom");
  const qTo = qs("qTo");
  if (qFrom && !qFrom.value) qFrom.value = new Date(Date.now() - 1000 * 60 * 60 * 24 * 30).toISOString().slice(0, 10);
  if (qTo && !qTo.value) qTo.value = new Date().toISOString().slice(0, 10);
  const costMonth = qs("costMonth");
  if (costMonth && !costMonth.value) costMonth.value = new Date().toISOString().slice(0, 7);
  const amlWeeklyEnd = qs("amlWeeklyEnd");
  if (amlWeeklyEnd && !amlWeeklyEnd.value) amlWeeklyEnd.value = defaultAmlWeeklyEndDate();
  loadStoreAllocations().catch(() => {});
  loadStockOnHandPage().catch(() => {});
  loadInventoryControl().catch(() => {});
  loadLocations().catch(() => {});
  loadStockBins().catch(() => {});
  loadStockDepth().catch(() => {});
  loadReplenishmentSuggestions().catch(() => {});
  loadCycleSessions().catch(() => {});
  loadLubeStockOnHand().catch(() => {});
  loadLubeReorderAlerts().catch(() => {});
  applyDefaultLocationsToInputs();
  loadBinCodeOptionsForLocation(qs("saLocation")?.value || "", "saBinCodeOptions").catch(() => {});
  loadBinCodeOptionsForLocation(qs("msLocation")?.value || "", "msBinCodeOptions").catch(() => {});
  updateManualStockCostRowVisibility();
  updateManualStockPartDesc();
  updateManualLubePartDesc();
  updateReceiveLubePartDesc();
  updateLubeMinPartDesc();
  loadLubeAnalytics().catch(() => {});
  loadLubeMappings().catch(() => {});
  setProcurementChainInputsFromConfig();
  updateProcurementChainPreview();
  setProcurementKpiFilter("all");
  const prJournalStart = qs("prJournalStart");
  const prJournalEnd = qs("prJournalEnd");
  if (prJournalStart && !prJournalStart.value) prJournalStart.value = new Date(Date.now() - 1000 * 60 * 60 * 24 * 30).toISOString().slice(0, 10);
  if (prJournalEnd && !prJournalEnd.value) prJournalEnd.value = new Date().toISOString().slice(0, 10);
  loadRequisitions().catch(() => {});
  loadPurchaseOrders().catch(() => {});
  loadProcurementExceptions().catch(() => {});
  loadOperations().catch(() => {});
  loadSiteZones().catch(() => {});
  loadSiteDailyEntries().catch(() => {});
  loadSiteTargets().catch(() => {});
  loadSiteDelays().catch(() => {});
  loadSiteDashboard().catch(() => {});
  loadDispatchTrips().catch(() => {});
  loadQualityCenter().catch(() => {});
  // Fuel endpoints are expensive on large datasets; keep initial page load responsive.
  // Users can load these manually from the Fuel tab buttons.
  if (getSessionRoles().some((r) => ["admin", "supervisor"].includes(r))) {
    loadCostSettings().catch(() => {});
  }
  loadLegalDepartments().catch(() => {});
  loadLegalDocs().catch(() => {});
  loadLegalExpiry().catch(() => {});
  if (getSessionRoles().some((r) => ["admin", "supervisor"].includes(r))) {
    loadAuditLogs().catch(() => {});
    loadApprovalRequests().catch(() => {});
    loadSmtpSettings().catch(() => {});
    loadPushNotificationSettings().catch(() => {});
    loadPdfReportSettings().catch(() => {});
    loadBackupFiles().catch(() => {});
  }
  loadCodePickers().catch(() => {});
  populateThresholdInputs();
  populateLdvPrestartThresholdInputs();
  loadDashboard().catch((e) => setStatus("Dashboard error: " + e.message));

  const legalList = qs("legalList");
  if (legalList) {
    legalList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const dl = target.getAttribute("data-legal-download-id");
      const ar = target.getAttribute("data-legal-archive-id");
      const active = target.getAttribute("data-legal-active");
      const stId = target.getAttribute("data-legal-status-id");
      const st = target.getAttribute("data-legal-status");
      const actionsId = target.getAttribute("data-legal-actions-id");
      if (dl) {
        downloadLegalDoc(dl);
        return;
      }
      if (actionsId) {
        showLegalActions(actionsId);
        return;
      }
      if (stId && st) {
        setLegalStatus(stId, st);
        return;
      }
      if (ar && active != null) {
        archiveLegalDoc(ar, Number(active) === 1);
      }
    });
  }

  const approvalList = qs("approvalList");
  if (approvalList) {
    approvalList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const aId = target.getAttribute("data-approval-approve-id");
      const rId = target.getAttribute("data-approval-reject-id");
      if (aId) {
        decideApprovalRequest(aId, "approve");
        return;
      }
      if (rId) {
        decideApprovalRequest(rId, "reject");
      }
    });
  }

  const approvalKpiStrip = qs("approvalKpiStrip");
  if (approvalKpiStrip) {
    approvalKpiStrip.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const btn = target.closest("[data-approval-kpi-filter]");
      if (!(btn instanceof HTMLElement)) return;
      const filter = btn.getAttribute("data-approval-kpi-filter");
      if (filter == null) return;
      const statusEl = qs("approvalStatus");
      if (statusEl) statusEl.value = filter;
      loadApprovalRequests().catch((e) => setStatus("Approvals error: " + e.message));
    });
  }

  const lubeAnalyticsList = qs("lubeAnalyticsList");
  if (lubeAnalyticsList) {
    lubeAnalyticsList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const btn = target.closest("[data-map-oil-key]");
      if (!(btn instanceof HTMLElement)) return;
      const oilKey = String(btn.getAttribute("data-map-oil-key") || "").trim();
      const partCode = String(btn.getAttribute("data-map-part-code") || "").trim();
      const oilEl = qs("lubeMapOilKey");
      const partEl = qs("lubeMapPartCode");
      if (oilEl) oilEl.value = oilKey;
      if (partEl) partEl.value = partCode;
      setStatus("Mapping fields pre-filled from selected lube row.");
    });
  }

  const procurementList = qs("procurementList");
  if (procurementList) {
    procurementList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const advanceId = target.getAttribute("data-pr-advance-id");
      const advanceStatus = target.getAttribute("data-pr-advance-status");
      const submitId = target.getAttribute("data-pr-submit-id");
      const finalizeId = target.getAttribute("data-pr-finalize-id");
      const postId = target.getAttribute("data-pr-post-id");
      const routeId = target.getAttribute("data-pr-route-id");
      const approveId = target.getAttribute("data-pr-approve-id");
      const receiveId = target.getAttribute("data-pr-receive-id");
      const receiveHalfId = target.getAttribute("data-pr-receive-half-id");
      const receiveFullId = target.getAttribute("data-pr-receive-full-id");
      const createPoId = target.getAttribute("data-pr-create-po-id");
      const outstanding = target.getAttribute("data-pr-outstanding");
      const duplicateJson = target.getAttribute("data-pr-duplicate");
      const openApprovalId = target.getAttribute("data-pr-open-approval-id");
      if (advanceId && advanceStatus) {
        advanceRequisitionStage(advanceId, advanceStatus).catch((e) => setStatus(`Advance failed: ${e.message || e}`));
        return;
      }
      if (finalizeId) {
        fetchJson(`${API}/api/procurement/requisitions/${finalizeId}/finalize`, { method: "POST", headers: { "Content-Type": "application/json" } })
          .then((res) => {
            setText("procurementResult", JSON.stringify(res, null, 2));
            return loadRequisitions();
          })
          .catch((e) => setStatus(`Finalize failed: ${e.message || e}`));
        return;
      }
      if (postId) {
        fetchJson(`${API}/api/procurement/requisitions/${postId}/post`, { method: "POST", headers: { "Content-Type": "application/json" } })
          .then((res) => {
            setText("procurementResult", JSON.stringify(res, null, 2));
            return loadRequisitions();
          })
          .catch((e) => setStatus(`Post failed: ${e.message || e}`));
        return;
      }
      if (routeId) {
        launchApprovalRouteForRequisition(routeId)
          .then(() => loadRequisitions())
          .catch((e) => setStatus(`Route failed: ${e.message || e}`));
        return;
      }
      if (approveId) {
        approveCurrentStepForRequisition(approveId)
          .then(() => loadRequisitions())
          .catch((e) => setStatus(`Approve failed: ${e.message || e}`));
        return;
      }
      if (submitId) {
        fetchJson(`${API}/api/procurement/requisitions/${submitId}/finalize`, { method: "POST", headers: { "Content-Type": "application/json" } })
          .then(() => fetchJson(`${API}/api/procurement/requisitions/${submitId}/post`, { method: "POST", headers: { "Content-Type": "application/json" } }))
          .then(() =>
            fetchJson(`${API}/api/procurement/requisitions/${submitId}/approvers`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ approvers: [{ name: "approver1" }] }),
            })
          )
          .then(() => fetchJson(`${API}/api/procurement/requisitions/${submitId}/send-approval`, { method: "POST", headers: { "Content-Type": "application/json" } }))
          .then((res) => {
            setText("procurementResult", JSON.stringify(res, null, 2));
            return loadRequisitions();
          })
          .catch((e) => setStatus(`Quick send failed: ${e.message || e}`));
        return;
      }
      if (receiveId) {
        requestRequisitionReceive(receiveId);
        return;
      }
      if (receiveHalfId) {
        requestRequisitionReceiveHalf(receiveHalfId, Number(outstanding || 0));
        return;
      }
      if (receiveFullId) {
        requestRequisitionReceiveFull(receiveFullId, Number(outstanding || 0));
        return;
      }
      if (createPoId) {
        createPoFromRequisition(createPoId).catch((e) => setStatus(`Create PO failed: ${e.message || e}`));
        return;
      }
      if (duplicateJson) {
        duplicateRequisitionFromRow(duplicateJson);
        return;
      }
      if (openApprovalId) {
        const statusEl = qs("approvalStatus");
        const moduleEl = qs("approvalModule");
        const actionEl = qs("approvalAction");
        if (statusEl) statusEl.value = "";
        if (moduleEl) moduleEl.value = "procurement";
        if (actionEl) actionEl.value = "";
        switchTab("approvals");
        loadApprovalRequests().catch(() => {});
        setStatus(`Showing approvals. Latest request id: #${openApprovalId}`);
      }
    });
  }

  const prPoList = qs("prPoList");
  if (prPoList) {
    prPoList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const openId = target.getAttribute("data-pr-po-open");
      const approveId = target.getAttribute("data-pr-po-approve");
      const sendId = target.getAttribute("data-pr-po-send");
      if (openId) {
        openPurchaseOrder(openId).catch((e) => setStatus(`PO detail error: ${e.message || e}`));
        return;
      }
      if (approveId) {
        approvePurchaseOrder(approveId).catch((e) => setStatus(`PO approve error: ${e.message || e}`));
        return;
      }
      if (sendId) {
        sendPurchaseOrder(sendId).catch((e) => setStatus(`PO send error: ${e.message || e}`));
      }
    });
  }

  const prExList = qs("prExceptionsList");
  if (prExList) {
    prExList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const resolveId = target.getAttribute("data-pr-ex-resolve");
      if (resolveId) {
        resolveProcurementException(resolveId).catch((e) => setStatus(`Exception resolve error: ${e.message || e}`));
      }
    });
  }

  const sbCycleSessionsList = qs("sbCycleSessionsList");
  if (sbCycleSessionsList) {
    sbCycleSessionsList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const submitId = target.getAttribute("data-sb-cs-submit");
      const approveId = target.getAttribute("data-sb-cs-approve");
      const countOneId = target.getAttribute("data-sb-cs-countone");
      if (submitId) {
        submitCycleSession(submitId).catch((e) => setStatus(`Cycle submit error: ${e.message || e}`));
        return;
      }
      if (approveId) {
        approveCycleSession(approveId).catch((e) => setStatus(`Cycle approve error: ${e.message || e}`));
        return;
      }
      if (countOneId) {
        addOnePartCountToSession(countOneId).catch((e) => setStatus(`Cycle line upsert error: ${e.message || e}`));
      }
    });
  }

  ["sfPlan", "sfReview", "sfRoute", "sfApprove", "sfPoReady", "sfReceive"].forEach((laneId) => {
    const lane = qs(laneId);
    if (!lane) return;
    lane.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const advanceId = target.getAttribute("data-pr-advance-id");
      const advanceStatus = target.getAttribute("data-pr-advance-status");
      if (!advanceId || !advanceStatus) return;
      advanceRequisitionStage(advanceId, advanceStatus).catch((e) => setStatus(`Advance failed: ${e.message || e}`));
    });
  });

  const dispatchList = qs("dispatchList");
  if (dispatchList) {
    dispatchList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const id = target.getAttribute("data-dp-status-id");
      const next = target.getAttribute("data-dp-next");
      if (!id || !next) return;
      updateDispatchTripStatus(id, next).catch((e) => setStatus(`Dispatch status update failed: ${e.message || e}`));
    });
  }
  const dispatchExceptionsList = qs("dispatchExceptionsList");
  if (dispatchExceptionsList) {
    dispatchExceptionsList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const id = target.getAttribute("data-dp-ex-id");
      const next = target.getAttribute("data-dp-ex-next");
      if (!id || !next) return;
      resolveDispatchException(id, next).catch((e) => setStatus(`Dispatch exception update failed: ${e.message || e}`));
    });
  }
  const qualityList = qs("qualityList");
  if (qualityList) {
    qualityList.addEventListener("click", (evt) => {
      const target = evt.target;
      if (!(target instanceof HTMLElement)) return;
      const resolveBtn = target.closest("[data-q-resolve]");
      if (resolveBtn instanceof HTMLElement) {
        const mode = resolveBtn.getAttribute("data-q-resolve");
        const entity = resolveBtn.getAttribute("data-q-entity");
        const date = resolveBtn.getAttribute("data-q-date");
        resolveQualityIssueNow(mode, entity, date).catch((e) => setStatus(`Quality resolve failed: ${e.message || e}`));
        return;
      }
      const btn = target.closest("[data-q-fix]");
      if (!(btn instanceof HTMLElement)) return;
      const type = btn.getAttribute("data-q-type");
      const asset = btn.getAttribute("data-q-asset");
      const entity = btn.getAttribute("data-q-entity");
      const date = btn.getAttribute("data-q-date");
      openQualityFix(type, asset, entity, date);
    });
  }

  ["sfCountPlan", "sfCountReview", "sfCountRoute", "sfCountApprove", "sfCountPoReady", "sfCountReceive"].forEach((id) => {
    const btn = qs(id);
    if (!btn) return;
    btn.addEventListener("click", () => {
      const statusEl = qs("prStatusFilter");
      if (id === "sfCountReceive") {
        if (statusEl) statusEl.value = "";
        setProcurementKpiFilter("receive_set");
      } else {
        const status = String(btn.getAttribute("data-sf-status") || "").trim();
        if (statusEl) statusEl.value = status;
        setProcurementKpiFilter("all");
      }
      loadRequisitions().catch((e) => setStatus("Requisition load error: " + e.message));
    });
  });

  loadDocHeaders().catch(() => {});
  loadDocDrafts().catch(() => {});
}

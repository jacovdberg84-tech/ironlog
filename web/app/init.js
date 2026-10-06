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

  // Feature wiring, in the original start-up order. Each function lives in its feature's file.
  wireDashboardControls();
  wireBorrisControls();
  wireDocumentControls();
  wireLubeControls();
  wireStockReceive();
  wireLubeModel();
  wireProcurementControls();
  wireFuelLogControls();
  wireSiteOpsControls();
  wireFuelReportControls();
  wireStockControls();
  wireAuditControls();
  wireLoginControls();
  wireAdminControls();
  wireApprovalControls();
  wireUploadControls();
  wireReportControls();
  wireBreakdownControls();
  wireInventoryControls();
  wireDailyInputControls();
  wireSafetyAdminControls();
  wireTelematicsControls();
  wireOfflineSync();
  wireAssetControls();
  loadStartupData();
  wireComplianceLists();
  wireLubeAnalyticsList();
  wireProcurementLists();
  wireCycleCountList();
  wireSupplyFlowLanes();
  wireSiteOpsLists();
  wireSupplyFlowCounters();
  loadDocumentsOnStartup();
}

/** Start-up: Default dates and the first data loads. Called once from init() in init.js. */
function loadStartupData() {
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
}

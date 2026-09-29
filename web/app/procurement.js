// IRONLOG/web/app/procurement.js — Requisitions, approval chains, purchase orders, journals.
// Part of the main app; index.html loads these files in order and they share one global scope.

async function createRequisition() {
  const payload = {
    part_code: (qs("prPartCode")?.value || "").trim(),
    qty_requested: Number(qs("prQty")?.value || 0),
    estimated_value: (qs("prValue")?.value || "").trim() === "" ? undefined : Number(qs("prValue")?.value || 0),
    needed_by_date: (qs("prNeedBy")?.value || "").trim() || undefined,
    bill_to: (qs("prBillTo")?.value || "workshop").trim(),
    request_type: (qs("prRequestType")?.value || "site").trim(),
    supplier_name: (qs("prSupplier")?.value || "").trim() || undefined,
    po_number: (qs("prPo")?.value || "").trim() || undefined,
    notes: (qs("prNotes")?.value || "").trim() || undefined,
  };
  setStatus("Creating requisition...");
  const res = await fetchJson(`${API}/api/procurement/requisitions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadRequisitions();
  setStatus("Requisition created.");
}

async function requestRequisitionApproval(id) {
  const reqId = Number(id || 0);
  if (!reqId) return;
  setStatus(`Submitting requisition #${reqId} for approval...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/request-approval`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadRequisitions().catch(() => {}), loadApprovalRequests().catch(() => {})]);
  setStatus(`Requisition #${reqId} sent for approval.`);
}

async function requestRequisitionReceive(id) {
  const reqId = Number(id || 0);
  if (!reqId) return;
  const qtyRaw = prompt("Receive quantity:");
  if (qtyRaw == null) return;
  const qty_receive = Number(qtyRaw);
  if (!Number.isFinite(qty_receive) || qty_receive <= 0) {
    alert("Receive quantity must be > 0.");
    return;
  }
  const reference = prompt("Reference (GRN/Invoice/PO):", `requisition:${reqId}`) || `requisition:${reqId}`;
  setStatus(`Submitting receive request for requisition #${reqId}...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/request-receive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ qty_receive, reference }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadRequisitions().catch(() => {}), loadApprovalRequests().catch(() => {})]);
  setStatus(`Receive request submitted for requisition #${reqId}.`);
}

async function requestRequisitionReceiveFull(id, qtyOutstanding) {
  const reqId = Number(id || 0);
  const qty_receive = Number(qtyOutstanding || 0);
  if (!reqId) return;
  if (!Number.isFinite(qty_receive) || qty_receive <= 0) {
    alert("No outstanding quantity to receive.");
    return;
  }
  const reference = prompt("Reference (GRN/Invoice/PO):", `requisition:${reqId}:full`) || `requisition:${reqId}:full`;
  setStatus(`Submitting full receive request for requisition #${reqId}...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/request-receive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ qty_receive, reference }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadRequisitions().catch(() => {}), loadApprovalRequests().catch(() => {})]);
  setStatus(`Full receive request submitted for requisition #${reqId}.`);
}

async function requestRequisitionReceiveHalf(id, qtyOutstanding) {
  const reqId = Number(id || 0);
  const outstanding = Number(qtyOutstanding || 0);
  if (!reqId) return;
  if (!Number.isFinite(outstanding) || outstanding <= 0) {
    alert("No outstanding quantity to receive.");
    return;
  }
  const qty_receive = Number((outstanding * 0.5).toFixed(2));
  const reference = prompt("Reference (GRN/Invoice/PO):", `requisition:${reqId}:half`) || `requisition:${reqId}:half`;
  setStatus(`Submitting 50% receive request for requisition #${reqId}...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/request-receive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ qty_receive, reference }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadRequisitions().catch(() => {}), loadApprovalRequests().catch(() => {})]);
  setStatus(`50% receive request submitted for requisition #${reqId}.`);
}

async function duplicateRequisitionFromRow(rowJson) {
  let row = null;
  try {
    row = JSON.parse(String(rowJson || "{}"));
  } catch {
    row = null;
  }
  if (!row || !row.part_code) return;

  const qtyDefault = Number(row.qty_requested || 1);
  const qtyRaw = prompt("Duplicate requisition quantity:", String(qtyDefault));
  if (qtyRaw == null) return;
  const qty_requested = Number(qtyRaw);
  if (!Number.isFinite(qty_requested) || qty_requested <= 0) {
    alert("Quantity must be > 0.");
    return;
  }
  const needBy = prompt("Needed by date (YYYY-MM-DD, optional):", String(row.needed_by_date || "")) || "";
  const payload = {
    part_code: String(row.part_code || "").trim(),
    qty_requested,
    needed_by_date: needBy.trim() || undefined,
    supplier_name: String(row.supplier_name || "").trim() || undefined,
    po_number: String(row.po_number || "").trim() || undefined,
    notes: `Duplicate of REQ #${row.id}${row.notes ? ` | ${row.notes}` : ""}`,
  };
  setStatus("Creating duplicate requisition...");
  const res = await fetchJson(`${API}/api/procurement/requisitions`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadRequisitions();
  setStatus(`Duplicate requisition created from REQ #${row.id}.`);
}

function getProcurementChainConfig() {
  const fallback = {
    tier1Max: 5000,
    tier1Chain: "supervisor",
    tier2Max: 25000,
    tier2Chain: "supervisor,manager",
    tier3Chain: "supervisor,manager,finance,admin",
  };
  try {
    const raw = localStorage.getItem("ironlog.procurement.chainConfig");
    if (!raw) return fallback;
    const parsed = JSON.parse(raw);
    return {
      tier1Max: Number(parsed?.tier1Max || fallback.tier1Max),
      tier1Chain: String(parsed?.tier1Chain || fallback.tier1Chain),
      tier2Max: Number(parsed?.tier2Max || fallback.tier2Max),
      tier2Chain: String(parsed?.tier2Chain || fallback.tier2Chain),
      tier3Chain: String(parsed?.tier3Chain || fallback.tier3Chain),
    };
  } catch {
    return fallback;
  }
}

function setProcurementChainInputsFromConfig() {
  const cfg = getProcurementChainConfig();
  if (qs("prTier1Max")) qs("prTier1Max").value = String(cfg.tier1Max);
  if (qs("prTier1Chain")) qs("prTier1Chain").value = cfg.tier1Chain;
  if (qs("prTier2Max")) qs("prTier2Max").value = String(cfg.tier2Max);
  if (qs("prTier2Chain")) qs("prTier2Chain").value = cfg.tier2Chain;
  if (qs("prTier3Chain")) qs("prTier3Chain").value = cfg.tier3Chain;
}

function saveProcurementChainConfig() {
  const tier1Max = Number(qs("prTier1Max")?.value || 0);
  const tier2Max = Number(qs("prTier2Max")?.value || 0);
  const tier1Chain = String(qs("prTier1Chain")?.value || "").trim();
  const tier2Chain = String(qs("prTier2Chain")?.value || "").trim();
  const tier3Chain = String(qs("prTier3Chain")?.value || "").trim();
  if (!Number.isFinite(tier1Max) || tier1Max < 0) throw new Error("Tier 1 max value is invalid.");
  if (!Number.isFinite(tier2Max) || tier2Max < 0) throw new Error("Tier 2 max value is invalid.");
  if (tier2Max < tier1Max) throw new Error("Tier 2 max must be greater than or equal to Tier 1 max.");
  if (!tier1Chain || !tier2Chain || !tier3Chain) throw new Error("All tier chains are required.");
  localStorage.setItem("ironlog.procurement.chainConfig", JSON.stringify({ tier1Max, tier1Chain, tier2Max, tier2Chain, tier3Chain }));
  setStatus("Approval chain rules saved.");
}

function pickApprovalChainForValue(value) {
  const cfg = getProcurementChainConfig();
  const v = Number(value || 0);
  if (!Number.isFinite(v) || v <= cfg.tier1Max) return cfg.tier1Chain;
  if (v <= cfg.tier2Max) return cfg.tier2Chain;
  return cfg.tier3Chain;
}

function procurementTierMetaByValue(value) {
  const cfg = getProcurementChainConfig();
  const v = Number(value || 0);
  if (!Number.isFinite(v) || v <= cfg.tier1Max) return { label: "TIER 1", cls: "tier-1" };
  if (v <= cfg.tier2Max) return { label: "TIER 2", cls: "tier-2" };
  return { label: "TIER 3", cls: "tier-3" };
}

function updateProcurementChainPreview() {
  const badge = qs("prChainTierBadge");
  const setBadge = (label, cls) => {
    if (!badge) return;
    badge.className = "pill tier-badge";
    if (cls) badge.classList.add(cls);
    badge.textContent = label;
  };
  const override = String(qs("prApproverChain")?.value || "").trim();
  if (override) {
    setBadge("MANUAL", "tier-manual");
    setText("prChainPreview", `Manual override: ${override}`);
    return;
  }
  const cfg = getProcurementChainConfig();
  const valueRaw = String(qs("prValue")?.value || "").trim();
  const valueNum = valueRaw === "" ? 0 : Number(valueRaw);
  const chain = pickApprovalChainForValue(valueNum);
  const displayValue = Number.isFinite(valueNum) ? valueNum.toFixed(2) : "0.00";
  const tier = procurementTierMetaByValue(valueNum);
  setBadge(tier.label, tier.cls);
  setText("prChainPreview", `Value R${displayValue} -> ${chain}`);
}

async function launchApprovalRouteForRequisition(reqId) {
  const row = procurementRowsCache.find((r) => Number(r.id) === Number(reqId));
  const estimatedValue = Number(row?.estimated_value || 0);
  const chainInput = qs("prApproverChain");
  const typedChain = String(chainInput?.value || "").trim();
  const valueBasedChain = pickApprovalChainForValue(estimatedValue);
  const namesRaw = typedChain || valueBasedChain || prompt("Approvers in order (comma separated names):", "approver1,approver2");
  if (!namesRaw) return;
  const approvers = namesRaw
    .split(",")
    .map((x) => String(x || "").trim())
    .filter(Boolean)
    .map((name) => ({ name }));
  await fetchJson(`${API}/api/procurement/requisitions/${reqId}/approvers`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ approvers }),
  });
  if (chainInput && typedChain) {
    chainInput.value = namesRaw;
  }
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/send-approval`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
}

async function approveCurrentStepForRequisition(reqId) {
  const who = prompt("Approver name:", getSessionUser()) || getSessionUser();
  const comment = prompt("Approval comment (optional):", "") || "";
  const res = await fetchJson(`${API}/api/procurement/requisitions/${reqId}/approve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ approver_name: who, comment }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
}

async function advanceRequisitionStage(reqId, currentStatus) {
  const id = Number(reqId || 0);
  const s = String(currentStatus || "").toLowerCase();
  if (!id || !s) return;
  setStatus(`Advancing requisition #${id} from ${s}...`);
  if (s === "draft") {
    const res = await fetchJson(`${API}/api/procurement/requisitions/${id}/finalize`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    setText("procurementResult", JSON.stringify(res, null, 2));
  } else if (s === "finalized") {
    const res = await fetchJson(`${API}/api/procurement/requisitions/${id}/post`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    setText("procurementResult", JSON.stringify(res, null, 2));
  } else if (s === "posted") {
    await launchApprovalRouteForRequisition(id);
  } else if (s === "approval_in_progress") {
    await approveCurrentStepForRequisition(id);
  } else {
    setStatus(`No advance action defined for status '${s}'.`);
    return;
  }
  await loadRequisitions();
  setStatus(`Requisition #${id} advanced.`);
}

function supplyFlowAdvanceButton(reqId, status) {
  const s = String(status || "").toLowerCase();
  const canAdvance = ["draft", "finalized", "posted", "approval_in_progress"].includes(s);
  if (canAdvance) {
    return `<button data-pr-advance-id="${reqId}" data-pr-advance-status="${s}" style="margin-top:8px;">Advance Stage</button>`;
  }
  let tip = "No advance action for this stage.";
  if (s === "approved_all") tip = "PO Ready is complete. Next action is Receive.";
  if (s === "approved") tip = "Receiving actions are available for this stage.";
  if (s === "received") tip = "Requisition fully received.";
  return `<button class="btn-disabled" title="${tip}" disabled style="margin-top:8px;">Advance Stage</button>`;
}

let procurementKpiFilter = "all";
let procurementRowsCache = [];

function setProcurementKpiFilter(filter) {
  procurementKpiFilter = String(filter || "all");
  qs("prKpiAll")?.classList.toggle("pill-active", procurementKpiFilter === "all");
  qs("prKpiApprovedOpen")?.classList.toggle("pill-active", procurementKpiFilter === "approved_open");
  qs("prKpiInFlow")?.classList.toggle("pill-active", procurementKpiFilter === "in_flow");
}

async function loadRequisitions() {
  const list = qs("procurementList");
  if (!list) return;
  const status = (qs("prStatusFilter")?.value || "").trim();
  const tierFilter = (qs("prTierFilter")?.value || "").trim().toLowerCase();
  const q = status ? `?status=${encodeURIComponent(status)}` : "";
  setStatus("Loading requisitions...");
  setSkeleton("procurementList", 2);
  const data = await fetchJson(`${API}/api/procurement/requisitions${q}`);
  const rows = Array.isArray(data.rows) ? data.rows : [];
  procurementRowsCache = rows;
  setText("prAllCount", rows.length);
  setText(
    "prApprovedOpenCount",
    rows.filter((r) => String(r.status || "").toLowerCase() === "approved" && Number(r.qty_outstanding || 0) > 0).length
  );
  setText(
    "prInFlowCount",
    rows.filter((r) => {
      const s = String(r.status || "").toLowerCase();
      return ["draft", "finalized", "posted", "approval_in_progress", "approved_all"].includes(s);
    }).length
  );
  renderSupplyFlowBoard(rows);
  let displayRows = rows;
  if (procurementKpiFilter === "approved_open") {
    displayRows = rows.filter((r) => String(r.status || "").toLowerCase() === "approved" && Number(r.qty_outstanding || 0) > 0);
  } else if (procurementKpiFilter === "in_flow") {
    displayRows = rows.filter((r) => {
      const s = String(r.status || "").toLowerCase();
      return ["draft", "finalized", "posted", "approval_in_progress", "approved_all"].includes(s);
    });
  } else if (procurementKpiFilter === "receive_set") {
    displayRows = rows.filter((r) => ["approved", "received"].includes(String(r.status || "").toLowerCase()));
  }
  if (tierFilter) {
    displayRows = displayRows.filter((r) => procurementTierMetaByValue(r.estimated_value).cls === tierFilter.replace("_", "-"));
  }
  list.innerHTML = "";
  displayRows.forEach((r) => {
    const s = String(r.status || "").toLowerCase();
    const stageLabel = supplyFlowStageLabel(s);
    const tier = procurementTierMetaByValue(r.estimated_value);
    const tierBadge = `<span class="pill tier-badge ${tier.cls}">${tier.label}</span>`;
    const finalizeBtn = s === "draft" ? `<button data-pr-finalize-id="${r.id}" style="margin-top:8px;">Finalize Request</button>` : "";
    const postBtn = s === "finalized" ? `<button data-pr-post-id="${r.id}" style="margin-top:8px;">Buyer Review & Post</button>` : "";
    const routeBtn = s === "posted" ? `<button data-pr-route-id="${r.id}" style="margin-top:8px;">Launch Approval Route</button>` : "";
    const approveBtn = s === "approval_in_progress" ? `<button data-pr-approve-id="${r.id}" style="margin-top:8px;">Approve Current Step</button>` : "";
    const submitBtn = s === "draft" ? `<button data-pr-submit-id="${r.id}" style="margin-top:8px;">Fast Track to Approval</button>` : "";
    const receiveBtn = ["approved", "approved_all"].includes(s) ? `<button data-pr-receive-id="${r.id}" style="margin-top:8px;">Request Receive</button>` : "";
    const createPoBtn = ["approved_all", "approved", "po_ready", "partially_received", "received"].includes(s)
      ? `<button data-pr-create-po-id="${r.id}" style="margin-top:8px;">Create / Open PO</button>`
      : "";
    const receiveHalfBtn = ["approved", "approved_all"].includes(s) && Number(r.qty_outstanding || 0) > 0
      ? `<button data-pr-receive-half-id="${r.id}" data-pr-outstanding="${Number(r.qty_outstanding || 0)}" style="margin-top:8px;">Receive 50%</button>`
      : "";
    const receiveFullBtn = ["approved", "approved_all"].includes(s) && Number(r.qty_outstanding || 0) > 0
      ? `<button data-pr-receive-full-id="${r.id}" data-pr-outstanding="${Number(r.qty_outstanding || 0)}" style="margin-top:8px;">Receive Full</button>`
      : "";
    const advanceBtn = supplyFlowAdvanceButton(r.id, s);
    const dupPayload = String(JSON.stringify({
      id: r.id,
      part_code: r.part_code,
      qty_requested: r.qty_requested,
      needed_by_date: r.needed_by_date,
      supplier_name: r.supplier_name,
      po_number: r.po_number,
      notes: r.notes,
    })).replace(/"/g, "&quot;");
    list.appendChild(
      item(
        `<b>REQ #${r.id}</b> <span class="pill blue">${stageLabel}</span> - ${r.part_code || "-"} (${r.part_name || "-"})` +
          `<br>${tierBadge}` +
          `<br><small>Site Req No: ${r.site_request_no || "-"} | Bill to: ${r.bill_to || "-"} | Type: ${r.request_type || "-"}</small>` +
          `<br><small>Requested: ${Number(r.qty_requested || 0).toFixed(1)} | Received: ${Number(r.qty_received || 0).toFixed(1)} | Outstanding: ${Number(r.qty_outstanding || 0).toFixed(1)} | Need by: ${r.needed_by_date || "-"}</small>` +
          `<br><small>Req Value: ${r.estimated_value == null ? "-" : Number(r.estimated_value).toFixed(2)}</small>` +
          `<br><small>Supplier: ${r.supplier_name || "-"} | PO: ${r.po_number || "-"}</small>` +
          `<br><small>Requester: ${r.requester || "-"} | ${r.created_at || "-"}</small>` +
          (r.latest_approval_id ? `<br><small>Approval: #${r.latest_approval_id} (${r.latest_approval_status || "-"})</small>` : "") +
          (r.notes ? `<br><small>${r.notes}</small>` : "") +
          `<br>${advanceBtn} ${finalizeBtn} ${postBtn} ${routeBtn} ${approveBtn} ${submitBtn} ${receiveBtn} ${receiveHalfBtn} ${receiveFullBtn} ${createPoBtn} <button data-pr-duplicate="${dupPayload}" style="margin-top:8px;">Duplicate</button> ${r.latest_approval_id ? `<button data-pr-open-approval-id="${r.latest_approval_id}" style="margin-top:8px;">Open Approval</button>` : ""}`
      )
    );
  });
  if (!displayRows.length) list.appendChild(item("<small>No requisitions found.</small>"));
  setStatus("Requisitions ready.");
}

function supplyFlowStageLabel(status) {
  const s = String(status || "").toLowerCase();
  if (s === "draft") return "Plan";
  if (s === "finalized") return "Review";
  if (s === "posted") return "Approval Route";
  if (s === "approval_in_progress") return "Approvals";
  if (s === "approved_all") return "PO Ready";
  if (s === "approved" || s === "received") return "Receive";
  return s || "Unknown";
}

function renderSupplyFlowBoard(rows) {
  const lanes = {
    plan: qs("sfPlan"),
    review: qs("sfReview"),
    route: qs("sfRoute"),
    approve: qs("sfApprove"),
    po_ready: qs("sfPoReady"),
    receive: qs("sfReceive"),
  };
  const laneCounts = {
    plan: 0,
    review: 0,
    route: 0,
    approve: 0,
    po_ready: 0,
    receive: 0,
  };
  Object.values(lanes).forEach((el) => {
    if (el) el.innerHTML = "";
  });

  const add = (laneKey, html) => {
    const lane = lanes[laneKey];
    if (lane) lane.appendChild(item(html));
  };

  (Array.isArray(rows) ? rows : []).slice(0, 80).forEach((r) => {
    const s = String(r.status || "").toLowerCase();
    const tier = procurementTierMetaByValue(r.estimated_value);
    const snippet =
      `<b>#${r.id}</b> ${r.part_code || "-"} <span class="pill tier-badge ${tier.cls}">${tier.label}</span><br><small>Out: ${Number(r.qty_outstanding || 0).toFixed(1)} | ${r.bill_to || "-"}</small>` +
      `<br>${supplyFlowAdvanceButton(r.id, s)}`;
    if (s === "draft") {
      laneCounts.plan += 1;
      add("plan", snippet);
    } else if (s === "finalized") {
      laneCounts.review += 1;
      add("review", snippet);
    } else if (s === "posted") {
      laneCounts.route += 1;
      add("route", snippet);
    } else if (s === "approval_in_progress") {
      laneCounts.approve += 1;
      add("approve", snippet);
    } else if (s === "approved_all") {
      laneCounts.po_ready += 1;
      add("po_ready", snippet);
    } else {
      laneCounts.receive += 1;
      add("receive", snippet);
    }
  });

  Object.entries(lanes).forEach(([k, el]) => {
    if (el && !el.children.length) {
      el.appendChild(item(`<small>No requests in ${k.replace("_", " ")}.</small>`));
    }
  });

  setText("sfCountPlan", String(laneCounts.plan));
  setText("sfCountReview", String(laneCounts.review));
  setText("sfCountRoute", String(laneCounts.route));
  setText("sfCountApprove", String(laneCounts.approve));
  setText("sfCountPoReady", String(laneCounts.po_ready));
  setText("sfCountReceive", String(laneCounts.receive));
}

let procurementLastJournalBatchId = "";

async function loadPurchaseOrders() {
  const list = qs("prPoList");
  if (!list) return;
  const status = String(qs("prPoStatusFilter")?.value || "").trim();
  const q = status ? `?status=${encodeURIComponent(status)}` : "";
  setStatus("Loading purchase orders...");
  list.innerHTML = "";
  const data = await fetchJson(`${API}/api/procurement/purchase-orders${q}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  rows.forEach((r) => {
    const poId = Number(r.id || 0);
    list.appendChild(item(
      `<b>PO #${poId}</b> <span class="pill blue">${String(r.status || "-")}</span> - ${r.po_number || "-"}`
      + `<br><small>Supplier: ${r.supplier_name || r.supplier_code || "-"}</small>`
      + `<br><small>Subtotal: ${Number(r.subtotal || 0).toFixed(2)} ${r.currency || "USD"} | Req: ${r.requisition_id || "-"}</small>`
      + `<br><button data-pr-po-open="${poId}" style="margin-top:8px;">Open PO</button>`
      + ` <button data-pr-po-approve="${poId}" style="margin-top:8px;">Approve</button>`
      + ` <button data-pr-po-send="${poId}" style="margin-top:8px;">Mark Sent</button>`
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No purchase orders found.</small>"));
  setStatus("Purchase orders ready.");
}

async function createPoFromRequisition(reqId) {
  const id = Number(reqId || 0);
  if (!id) return;
  setStatus(`Creating PO from requisition #${id}...`);
  const res = await fetchJson(`${API}/api/procurement/requisitions/${id}/create-po`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  if (res?.po_id) {
    const poEl = qs("prActivePoId");
    if (poEl) poEl.value = String(res.po_id);
  }
  await Promise.all([loadRequisitions().catch(() => {}), loadPurchaseOrders().catch(() => {})]);
  setStatus(`PO created from requisition #${id}.`);
}

async function openPurchaseOrder(poId) {
  const id = Number(poId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${id}/detail`);
  const po = res?.po || {};
  const lines = Array.isArray(res?.lines) ? res.lines : [];
  const receipts = Array.isArray(res?.receipts) ? res.receipts : [];
  setText("procurementResult", JSON.stringify(res, null, 2));
  const poEl = qs("prActivePoId");
  if (poEl) poEl.value = String(id);
  const recJson = qs("prReceiveLinesJson");
  if (recJson) {
    const sample = lines
      .filter((l) => Number(l.quantity_ordered || 0) > Number(l.quantity_received || 0))
      .slice(0, 3)
      .map((l) => ({ po_line_id: Number(l.id), quantity_received: Number((Number(l.quantity_ordered || 0) - Number(l.quantity_received || 0)).toFixed(2)) }));
    recJson.value = JSON.stringify(sample.length ? sample : [{ po_line_id: Number(lines[0]?.id || 0), quantity_received: 1 }]);
  }
  const invJson = qs("prInvoiceLinesJson");
  if (invJson) {
    const sample = lines.slice(0, 3).map((l) => ({
      po_line_id: Number(l.id || 0),
      quantity_invoiced: Number((Number(l.quantity_received || l.quantity_ordered || 0)).toFixed(2)),
      unit_price: Number(l.unit_price || 0),
    }));
    invJson.value = JSON.stringify(sample.length ? sample : [{ po_line_id: Number(lines[0]?.id || 0), quantity_invoiced: 1, unit_price: 0 }]);
  }
  setStatus(`PO ${po.po_number || id} loaded (${lines.length} lines, ${receipts.length} receipts).`);
}

async function approvePurchaseOrder(poId) {
  const id = Number(poId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${id}/approve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadPurchaseOrders();
  setStatus(`PO #${id} approved.`);
}

async function sendPurchaseOrder(poId) {
  const id = Number(poId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${id}/send`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadPurchaseOrders();
  setStatus(`PO #${id} marked sent.`);
}

async function postPoReceipt() {
  const poId = Number(qs("prActivePoId")?.value || 0);
  if (!poId) throw new Error("PO ID is required.");
  const location_code = String(qs("prReceiveLocationCode")?.value || "MAIN").trim().toUpperCase() || "MAIN";
  const raw = String(qs("prReceiveLinesJson")?.value || "").trim();
  if (!raw) throw new Error("Receipt lines JSON is required.");
  let lines = [];
  try {
    lines = JSON.parse(raw);
  } catch {
    throw new Error("Receipt lines JSON is invalid.");
  }
  const payload = { location_code, lines };
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${poId}/receive`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await Promise.all([loadPurchaseOrders().catch(() => {}), loadRequisitions().catch(() => {})]);
  setStatus(`Receipt posted for PO #${poId}.`);
}

async function capturePoInvoice() {
  const poId = Number(qs("prActivePoId")?.value || 0);
  if (!poId) throw new Error("PO ID is required.");
  const invoice_number = String(qs("prInvoiceNumber")?.value || "").trim();
  if (!invoice_number) throw new Error("Invoice number is required.");
  const raw = String(qs("prInvoiceLinesJson")?.value || "").trim();
  if (!raw) throw new Error("Invoice lines JSON is required.");
  let lines = [];
  try {
    lines = JSON.parse(raw);
  } catch {
    throw new Error("Invoice lines JSON is invalid.");
  }
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${poId}/invoices`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ invoice_number, lines }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  setStatus(`Invoice captured for PO #${poId}.`);
}

async function runPoThreeWayMatch() {
  const poId = Number(qs("prActivePoId")?.value || 0);
  if (!poId) throw new Error("PO ID is required.");
  const res = await fetchJson(`${API}/api/procurement/purchase-orders/${poId}/three-way-match`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ quantity_tolerance: 0, price_tolerance_pct: 5, total_tolerance: 1 }),
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadProcurementExceptions();
  setStatus(`3-way match complete for PO #${poId}.`);
}

async function loadProcurementExceptions() {
  const list = qs("prExceptionsList");
  if (!list) return;
  const status = String(qs("prExceptionStatus")?.value || "open").trim().toLowerCase();
  list.innerHTML = "";
  const data = await fetchJson(`${API}/api/procurement/exceptions?status=${encodeURIComponent(status)}`);
  const rows = Array.isArray(data?.rows) ? data.rows : [];
  rows.forEach((r) => {
    list.appendChild(item(
      `<b>EX #${Number(r.id || 0)}</b> <span class="pill ${String(r.severity || "").toLowerCase() === "high" ? "red" : "orange"}">${r.severity || "-"}</span> <span class="pill blue">${r.status || "-"}</span>`
      + `<br><small>PO: ${r.po_number || r.po_id || "-"} | Type: ${r.exception_type || "-"}</small>`
      + `<br><small>${r.details ? JSON.stringify(r.details) : "-"}</small>`
      + (String(r.status || "").toLowerCase() === "open" ? `<br><button data-pr-ex-resolve="${Number(r.id || 0)}" style="margin-top:8px;">Resolve</button>` : "")
    ));
  });
  if (!rows.length) list.appendChild(item("<small>No exceptions found.</small>"));
}

async function resolveProcurementException(exId) {
  const id = Number(exId || 0);
  if (!id) return;
  const res = await fetchJson(`${API}/api/procurement/exceptions/${id}/resolve`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
  });
  setText("procurementResult", JSON.stringify(res, null, 2));
  await loadProcurementExceptions();
  setStatus(`Exception #${id} resolved.`);
}

async function buildProcurementJournals() {
  const start = String(qs("prJournalStart")?.value || "").trim();
  const end = String(qs("prJournalEnd")?.value || "").trim();
  if (!start || !end) throw new Error("Journal start and end dates are required.");
  const batch_id = String(qs("prJournalBatchId")?.value || "").trim() || undefined;
  const default_cost_center_code = String(qs("prJournalCc")?.value || "").trim() || undefined;
  const res = await fetchJson(`${API}/api/procurement/journals/build`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ start, end, batch_id, default_cost_center_code }),
  });
  procurementLastJournalBatchId = String(res?.batch_id || "");
  setText("prJournalLastBatch", procurementLastJournalBatchId || "-");
  setText("procurementResult", JSON.stringify(res, null, 2));
  setStatus(`Journal batch built: ${procurementLastJournalBatchId || "-"}.`);
}

function currentProcurementJournalBatch() {
  const typed = String(qs("prJournalBatchId")?.value || "").trim();
  return typed || procurementLastJournalBatchId;
}

function exportProcurementJournalsCsv() {
  const batch = currentProcurementJournalBatch();
  if (!batch) throw new Error("Build journals first or enter batch ID.");
  return downloadAuthedFile(
    `${API}/api/procurement/journals/export.csv?batch_id=${encodeURIComponent(batch)}`,
    `IRONLOG_Procurement_Journals_${batch}.csv`,
  );
}

function exportProcurementJournalsXlsx() {
  const batch = currentProcurementJournalBatch();
  if (!batch) throw new Error("Build journals first or enter batch ID.");
  return downloadAuthedFile(
    `${API}/api/procurement/journals/export.xlsx?batch_id=${encodeURIComponent(batch)}`,
    `IRONLOG_Procurement_Journals_${batch}.xlsx`,
  );
}

/** Start-up: Requisition, purchase order and journal controls. Called once from init() in init.js. */
function wireProcurementControls() {
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
}

/** Start-up: Requisition, purchase order and exception list actions. Called once from init() in init.js. */
function wireProcurementLists() {
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
}

/** Start-up: Supply flow board lane actions. Called once from init() in init.js. */
function wireSupplyFlowLanes() {
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
}

/** Start-up: Supply flow counter buttons. Called once from init() in init.js. */
function wireSupplyFlowCounters() {
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
}

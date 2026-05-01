#!/usr/bin/env node

const API_BASE = process.env.API_BASE || "http://localhost:3001";
const SITE_CODE = process.env.SITE_CODE || "main";
const PART_CODE = process.env.PART_CODE || "TEST-PART";

function headers() {
  return {
    "Content-Type": "application/json",
    "x-user-role": "admin",
    "x-user-name": "smoke-runner",
    "x-user-roles": "admin,supervisor,stores,procurement",
    "x-site-code": SITE_CODE,
  };
}

async function call(method, path, body) {
  const res = await fetch(`${API_BASE}${path}`, {
    method,
    headers: headers(),
    body: body ? JSON.stringify(body) : undefined,
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg = json?.error || `HTTP ${res.status}`;
    throw new Error(`${method} ${path} failed: ${msg}`);
  }
  return json;
}

async function ensurePartExists() {
  // Uses stock movement create_if_missing path to avoid direct DB writes.
  await call("POST", "/api/stock/movement", {
    part_code: PART_CODE,
    part_name: "Smoke Test Part",
    movement_type: "in",
    quantity: 5,
    reference: "smoke-part-seed",
    create_if_missing: true,
    unit_cost: 1,
    cost_currency: "USD",
  });
}

async function run() {
  console.log(`API_BASE=${API_BASE} SITE_CODE=${SITE_CODE} PART_CODE=${PART_CODE}`);
  await ensurePartExists();

  console.log("1) Create supplier");
  await call("POST", "/api/procurement/suppliers", {
    supplier_code: "SMOKE-SUP",
    supplier_name: "Smoke Supplier",
    lead_time_days: 3,
    currency: "USD",
  });

  console.log("2) Create requisition");
  const req = await call("POST", "/api/procurement/requisitions", {
    part_code: PART_CODE,
    qty_requested: 2,
    estimated_value: 20,
    supplier_name: "Smoke Supplier",
    notes: "procurement smoke run",
  });
  const reqId = Number(req.id || 0);
  if (!reqId) throw new Error("Requisition ID missing");

  console.log("3) Move requisition to approved_all");
  await call("POST", `/api/procurement/requisitions/${reqId}/finalize`);
  await call("POST", `/api/procurement/requisitions/${reqId}/post`);
  await call("POST", `/api/procurement/requisitions/${reqId}/approvers`, {
    approvers: [{ name: "smoke-runner" }],
  });
  await call("POST", `/api/procurement/requisitions/${reqId}/send-approval`);
  await call("POST", `/api/procurement/requisitions/${reqId}/approve`, {
    approver_name: "smoke-runner",
    comment: "approved by smoke",
  });

  console.log("4) Create PO");
  const poCreate = await call("POST", `/api/procurement/requisitions/${reqId}/create-po`);
  const poId = Number(poCreate.po_id || 0);
  if (!poId) throw new Error("PO ID missing");

  console.log("5) Approve + send PO");
  await call("POST", `/api/procurement/purchase-orders/${poId}/approve`);
  await call("POST", `/api/procurement/purchase-orders/${poId}/send`);

  console.log("6) Load PO detail");
  const poDetail = await call("GET", `/api/procurement/purchase-orders/${poId}/detail`);
  const lines = Array.isArray(poDetail.lines) ? poDetail.lines : [];
  if (!lines.length) throw new Error("PO has no lines");
  const firstLine = lines[0];

  console.log("7) Receive PO");
  await call("POST", `/api/procurement/purchase-orders/${poId}/receive`, {
    location_code: "MAIN",
    lines: [{ po_line_id: Number(firstLine.id), quantity_received: 1 }],
  });

  console.log("8) Capture invoice");
  const invNumber = `SMOKE-INV-${Date.now()}`;
  await call("POST", `/api/procurement/purchase-orders/${poId}/invoices`, {
    invoice_number: invNumber,
    lines: [{ po_line_id: Number(firstLine.id), quantity_invoiced: 1, unit_price: Number(firstLine.unit_price || 0) }],
  });

  console.log("9) Run 3-way match");
  const match = await call("POST", `/api/procurement/purchase-orders/${poId}/three-way-match`, {
    quantity_tolerance: 0,
    price_tolerance_pct: 5,
    total_tolerance: 1,
  });

  console.log("10) Build journals");
  const today = new Date().toISOString().slice(0, 10);
  const journals = await call("POST", "/api/procurement/journals/build", {
    start: today,
    end: today,
    default_cost_center_code: "SMOKE-CC",
  });

  console.log("DONE");
  console.log(JSON.stringify({
    requisition_id: reqId,
    po_id: poId,
    match_exceptions: Number(match.exception_count || 0),
    journal_batch: journals.batch_id || null,
    journal_lines: Number(journals.lines || 0),
    balanced: Boolean(journals.balanced),
  }, null, 2));
}

run().catch((err) => {
  console.error("SMOKE FAILED:", err.message || err);
  process.exit(1);
});


// IRONLOG/api/utils/artisanInspection.js — the artisan (technician) inspection.
//
// A technician inspects a machine from the portal: every check is OK, Fault or
// N/A (a fault needs a comment and may have photos), plus the inspection
// details (type, shift, hour meter, where the machine is, overall result).
// Faults open one work order for the machine and put up the workshop notice the
// next operator sees on their pre-start. The checklist is English + Portuguese.
//
// Rows go in the existing artisan_inspections table (checklist_json keeps
// { key, label, ok, note } so older screens and reports still read it).

import fs from "node:fs";
import { buildPdfBuffer, ensurePageSpace, kvGrid, pdfBodyTop, sectionTitle } from "./pdfGenerator.js";
import { resolveStorageAbs } from "./storagePaths.js";
import { autoNoticeForFaults } from "./machineNotices.js";

export const ARTISAN_SECTIONS = [
  {
    id: "engine", en: "Engine & fluids", pt: "Motor e fluidos",
    items: [
      { key: "engine_oil", en: "Engine oil level and leaks", pt: "Nível de óleo do motor e fugas" },
      { key: "coolant", en: "Coolant level, radiator and hoses", pt: "Nível do líquido de arrefecimento, radiador e mangueiras" },
      { key: "fuel_system", en: "Fuel system and leaks", pt: "Sistema de combustível e fugas" },
      { key: "air_intake", en: "Air intake and filter indicator", pt: "Admissão de ar e indicador do filtro" },
      { key: "belts", en: "Belts and pulleys", pt: "Correias e polias" },
      { key: "lubrication", en: "Lubrication points / levels", pt: "Pontos de lubrificação / níveis" },
    ],
  },
  {
    id: "hydraulics", en: "Hydraulics", pt: "Hidráulica",
    items: [
      { key: "hydraulics", en: "Hydraulic hoses, leaks, and fittings", pt: "Mangueiras hidráulicas, fugas e conexões" },
      { key: "hyd_oil", en: "Hydraulic oil level", pt: "Nível do óleo hidráulico" },
      { key: "cylinders", en: "Cylinders, rams and seals", pt: "Cilindros, hastes e vedantes" },
    ],
  },
  {
    id: "electrical", en: "Electrical", pt: "Eléctrico",
    items: [
      { key: "electrical", en: "Electrical panels / cabling / lights", pt: "Painéis eléctricos / cablagem / luzes" },
      { key: "battery", en: "Battery, terminals and charging", pt: "Bateria, terminais e carregamento" },
      { key: "gauges", en: "Gauges and warning lights on the dash", pt: "Mostradores e luzes de aviso no painel" },
    ],
  },
  {
    id: "drive", en: "Brakes, steering & drive", pt: "Travões, direcção e transmissão",
    items: [
      { key: "brakes_steering", en: "Brakes / steering / controls response", pt: "Travões / direcção / resposta dos comandos" },
      { key: "transmission", en: "Transmission and drive", pt: "Transmissão e tracção" },
      { key: "tyres_tracks", en: "Tyres / tracks and wheel nuts", pt: "Pneus / rastos e porcas das rodas" },
    ],
  },
  {
    id: "structure", en: "Structure & attachments", pt: "Estrutura e acessórios",
    items: [
      { key: "structure", en: "Frame, boom and body (cracks or damage)", pt: "Chassis, lança e carroçaria (fissuras ou danos)" },
      { key: "attachments", en: "Bucket / GET / attachments, pins and bushes", pt: "Balde / dentes / acessórios, pinos e casquilhos" },
    ],
  },
  {
    id: "safety", en: "Cab & safety", pt: "Cabine e segurança",
    items: [
      { key: "prestart", en: "Pre-start visual condition (machine / plant)", pt: "Estado visual pré-arranque (máquina / instalação)" },
      { key: "guards", en: "Guards, covers, and safety devices", pt: "Protecções, tampas e dispositivos de segurança" },
      { key: "alarms", en: "Alarms, horn, and warning systems", pt: "Alarmes, buzina e sistemas de aviso" },
      { key: "seat_belt", en: "Seat and seat belt", pt: "Banco e cinto de segurança" },
      { key: "visibility", en: "Mirrors, windows and wipers", pt: "Espelhos, vidros e limpa-pára-brisas" },
      { key: "access", en: "Steps, handrails and access", pt: "Degraus, corrimãos e acesso" },
      { key: "fire", en: "Fire extinguisher / suppression system", pt: "Extintor / sistema de supressão de incêndio" },
    ],
  },
  {
    id: "area", en: "Housekeeping", pt: "Arrumação",
    items: [
      { key: "housekeeping", en: "Housekeeping around machine / plant", pt: "Arrumação à volta da máquina / instalação" },
    ],
  },
];

export const INSPECTION_TYPES = [
  { key: "daily", en: "Daily", pt: "Diária" },
  { key: "weekly", en: "Weekly", pt: "Semanal" },
  { key: "before_service", en: "Before service", pt: "Antes do serviço" },
  { key: "after_repair", en: "After repair", pt: "Depois da reparação" },
  { key: "other", en: "Other", pt: "Outra" },
];

export const OVERALL_RESULTS = [
  { key: "fit", en: "Fit for work", pt: "Apta para trabalhar" },
  { key: "restricted", en: "Fit with restrictions", pt: "Apta com restrições" },
  { key: "not_fit", en: "Not fit for work", pt: "Não apta para trabalhar" },
];

const ITEMS = ARTISAN_SECTIONS.flatMap((s) => s.items.map((i) => ({ ...i, section: s.id })));
const label = (list, key) => list.find((x) => x.key === key)?.en || key || "";

export function artisanTemplate() {
  return { sections: ARTISAN_SECTIONS, inspection_types: INSPECTION_TYPES, results: OVERALL_RESULTS };
}

function hasColumn(db, table, col) {
  return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
}

export function ensureArtisanSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS artisan_inspections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      asset_id INTEGER NOT NULL,
      uuid TEXT UNIQUE,
      site_code TEXT DEFAULT 'main',
      inspection_date TEXT NOT NULL,
      inspector_name TEXT,
      shift TEXT,
      notes TEXT,
      machine_hours REAL,
      live_hours_snapshot REAL,
      live_hours_source TEXT,
      checklist_json TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS artisan_inspection_photos (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      inspection_id INTEGER NOT NULL,
      item_key TEXT,
      file_path TEXT NOT NULL,
      caption TEXT,
      username TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_artisan_photos_insp ON artisan_inspection_photos(inspection_id);
  `);
  const cols = {
    form_number: "TEXT",
    inspection_type: "TEXT",
    location: "TEXT",
    overall_result: "TEXT",
    work_order_id: "INTEGER",
    inspector_username: "TEXT",
    client_id: "TEXT",
  };
  for (const [c, type] of Object.entries(cols)) {
    if (!hasColumn(db, "artisan_inspections", c)) db.prepare(`ALTER TABLE artisan_inspections ADD COLUMN ${c} ${type}`).run();
  }
}

export function formNumber(date = new Date()) {
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}${pad(date.getHours())}${pad(date.getMinutes())}`;
  return `AI-${stamp}-${Math.floor(Math.random() * 900 + 100)}`;
}

const fail = (message) => Object.assign(new Error(message), { status: 400 });

/**
 * answers: { key: "ok" | "fault" | "na" }, notes: { key: comment }.
 * Every check needs an answer and every fault a comment.
 */
export function readChecklist(answers = {}, notes = {}) {
  const missing = [];
  const checklist = ITEMS.map((i) => {
    const a = String(answers?.[i.key] || "").toLowerCase();
    const note = String(notes?.[i.key] || "").trim().slice(0, 500) || null;
    if (!["ok", "fault", "na"].includes(a)) missing.push(i.en);
    return { key: i.key, label: i.en, label_pt: i.pt, section: i.section, status: a || null, ok: a === "ok" ? true : a === "fault" ? false : null, note };
  });
  if (missing.length) throw fail(`Answer every check (${missing.length} left: ${missing.slice(0, 3).join(", ")}${missing.length > 3 ? "…" : ""}).`);
  const noComment = checklist.filter((c) => c.status === "fault" && !c.note);
  if (noComment.length) throw fail(`Say what is wrong for each fault (${noComment.map((c) => c.label).join(", ")}).`);
  return checklist;
}

export function faultsOf(checklist) {
  return checklist.filter((c) => c.ok === false).map((c) => ({ key: c.key, label: c.label, label_pt: c.label_pt || null, comment: c.note || "" }));
}

/**
 * Saves the inspection; faults open one work order (source "artisan_inspection")
 * and put up the operator notice. Returns { id, form_number, work_order_id, faults, fault_list }.
 */
export function saveArtisanInspection(db, input) {
  ensureArtisanSchema(db);
  // The same form sent again from the phone (e.g. after a failed photo) is saved once.
  const clientId = String(input.clientId || "").slice(0, 80) || null;
  const prior = clientId && db.prepare(`
    SELECT ai.id, ai.form_number, ai.work_order_id, ai.checklist_json, a.asset_code
    FROM artisan_inspections ai JOIN assets a ON a.id = ai.asset_id WHERE ai.client_id = ?
  `).get(clientId);
  if (prior) {
    let list = [];
    try { list = JSON.parse(prior.checklist_json || "[]"); } catch { list = []; }
    return { id: prior.id, form_number: prior.form_number, work_order_id: prior.work_order_id || null, faults: list.filter((c) => c.ok === false).length, asset_code: prior.asset_code, fault_list: [], duplicate: true };
  }
  const asset = db.prepare(`SELECT id, asset_code, asset_name FROM assets WHERE UPPER(asset_code) = UPPER(?)`).get(String(input.assetCode || "").trim());
  if (!asset) throw Object.assign(new Error("Machine not found"), { status: 404 });
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(input.date || "")) ? input.date : new Date().toISOString().slice(0, 10);
  const type = INSPECTION_TYPES.some((t) => t.key === input.type) ? input.type : "daily";
  const shift = ["day", "night"].includes(input.shift) ? input.shift : null;
  const result = OVERALL_RESULTS.some((r) => r.key === input.result) ? input.result : null;
  if (!result) throw fail("Choose the overall result: fit for work, fit with restrictions or not fit.");
  let hours = null;
  if (input.hours != null && String(input.hours).trim() !== "") {
    hours = Number(input.hours);
    if (!Number.isFinite(hours) || hours < 0) throw fail("Hour meter must be a number.");
  }
  const checklist = readChecklist(input.answers, input.comments);
  const faults = faultsOf(checklist);
  if (result === "not_fit" && !faults.length && !String(input.notes || "").trim()) {
    throw fail("Say why the machine is not fit for work (mark the fault or write a note).");
  }
  const notes = String(input.notes || "").trim().slice(0, 2000) || null;
  const location = String(input.location || "").trim().slice(0, 120) || null;
  const site = String(input.site || "main");
  const form = formNumber();

  const tx = db.transaction(() => {
    const id = Number(db.prepare(`
      INSERT INTO artisan_inspections (
        asset_id, uuid, site_code, inspection_date, inspector_name, inspector_username, form_number, shift, notes,
        machine_hours, live_hours_snapshot, live_hours_source, checklist_json, inspection_type, location, overall_result, client_id, updated_at
      ) VALUES (?, lower(hex(randomblob(16))), ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
    `).run(
      asset.id, site, date, input.inspector || null, input.username || null, form, shift, notes,
      hours, input.liveHours ?? null, input.liveSource || null, JSON.stringify(checklist), type, location, result,
      clientId,
    ).lastInsertRowid);

    let woId = null;
    if (faults.length) {
      const who = input.inspector || input.username || "technician";
      const description = [
        `Inspection fault${faults.length === 1 ? "" : "s"} found by ${who} on ${date} (${form})${result === "not_fit" ? " — machine NOT fit for work" : result === "restricted" ? " — fit with restrictions" : ""}:`,
        ...faults.map((f) => `- ${f.label}: ${f.comment}`),
      ].join("\n");
      const cols = ["asset_id", "source", "reference_id", "status"];
      const vals = [asset.id, "artisan_inspection", id, "open"];
      for (const [c, v] of [["site_code", site], ["job_description", description], ["priority", result === "not_fit" ? "high" : null]]) {
        if (v != null && hasColumn(db, "work_orders", c)) { cols.push(c); vals.push(v); }
      }
      woId = Number(db.prepare(`INSERT INTO work_orders (${cols.join(", ")}) VALUES (${cols.map(() => "?").join(", ")})`).run(...vals).lastInsertRowid);
      db.prepare(`UPDATE artisan_inspections SET work_order_id = ? WHERE id = ?`).run(woId, id);
      autoNoticeForFaults(db, { assetId: asset.id, assetCode: asset.asset_code, workOrderId: woId, faults, origin: "inspection", notFit: result === "not_fit" });
    }
    return { id, form_number: form, work_order_id: woId, faults: faults.length, asset_code: asset.asset_code, fault_list: faults };
  });
  return tx();
}

/** Keeps a photo on the inspection; on a fault work order it shows with the job too. */
export function addInspectionPhoto(db, { inspectionId, itemKey = null, relPath, caption = null, username = null }) {
  ensureArtisanSchema(db);
  const insp = db.prepare(`SELECT id, work_order_id FROM artisan_inspections WHERE id = ?`).get(Number(inspectionId));
  if (!insp) throw Object.assign(new Error("Inspection not found"), { status: 404 });
  const key = ITEMS.some((i) => i.key === itemKey) ? itemKey : null;
  const text = String(caption || "").trim().slice(0, 200) || (key ? label(ITEMS, key) : "Machine photo");
  const id = Number(db.prepare(`INSERT INTO artisan_inspection_photos (inspection_id, item_key, file_path, caption, username) VALUES (?, ?, ?, ?, ?)`)
    .run(insp.id, key, relPath, text, username).lastInsertRowid);
  const woPhotos = db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'work_order_photos'`).get();
  if (insp.work_order_id && woPhotos) {
    db.prepare(`INSERT INTO work_order_photos (work_order_id, username, file_path, caption, at) VALUES (?, ?, ?, ?, ?)`)
      .run(insp.work_order_id, username, relPath, `Inspection: ${text}`, new Date().toISOString());
  }
  return { id, inspection_id: insp.id, url: `/${relPath}` };
}

export function getArtisanInspection(db, id) {
  ensureArtisanSchema(db);
  const r = db.prepare(`
    SELECT ai.*, a.asset_code, a.asset_name, a.category
    FROM artisan_inspections ai JOIN assets a ON a.id = ai.asset_id WHERE ai.id = ?
  `).get(Number(id));
  if (!r) return null;
  let checklist = [];
  try { checklist = JSON.parse(r.checklist_json || "[]"); } catch { checklist = []; }
  const photos = db.prepare(`SELECT id, item_key, file_path, caption, username, created_at FROM artisan_inspection_photos WHERE inspection_id = ? ORDER BY id`).all(r.id);
  return { ...r, checklist: Array.isArray(checklist) ? checklist : [], photos };
}

/** Recent inspections (for the portal and the machine view). */
export function recentInspections(db, { assetId = null, username = null, limit = 20 } = {}) {
  ensureArtisanSchema(db);
  const where = [];
  const args = [];
  if (assetId) { where.push("ai.asset_id = ?"); args.push(Number(assetId)); }
  if (username) { where.push("LOWER(ai.inspector_username) = LOWER(?)"); args.push(username); }
  return db.prepare(`
    SELECT ai.id, ai.inspection_date, ai.inspector_name, ai.inspection_type, ai.overall_result, ai.work_order_id, ai.form_number,
      ai.checklist_json, a.asset_code, a.asset_name,
      (SELECT COUNT(*) FROM artisan_inspection_photos p WHERE p.inspection_id = ai.id) AS photo_count
    FROM artisan_inspections ai JOIN assets a ON a.id = ai.asset_id
    ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
    ORDER BY ai.inspection_date DESC, ai.id DESC LIMIT ?
  `).all(...args, Number(limit)).map(({ checklist_json, ...r }) => {
    let list = [];
    try { list = JSON.parse(checklist_json || "[]"); } catch { list = []; }
    return { ...r, faults: Array.isArray(list) ? list.filter((c) => c.ok === false).map((c) => c.label) : [] };
  });
}

const STATUS_TEXT = { ok: "OK", fault: "FAULT", na: "N/A" };

/** A4 report: details, checklist by section (faults in red), photos. */
export function buildArtisanInspectionPdf(insp) {
  return buildPdfBuffer((doc) => {
    const left = doc.page.margins.left;
    const width = doc.page.width - left - doc.page.margins.right;
    // The header is drawn on every page afterwards; start the text below it.
    const top = () => pdfBodyTop(doc, { siteName: "site" });
    doc.on("pageAdded", () => { doc.y = top(); });
    doc.y = Math.max(doc.y, top());
    sectionTitle(doc, "Artisan Inspection / Inspecção do Artífice");
    kvGrid(doc, [
      { k: "Inspection #", v: insp.id },
      { k: "Form No.", v: insp.form_number || "—" },
      { k: "Date", v: insp.inspection_date || "" },
      { k: "Type", v: insp.inspection_type ? label(INSPECTION_TYPES, insp.inspection_type) : "—" },
      { k: "Shift", v: insp.shift ? String(insp.shift).toUpperCase() : "—" },
      { k: "Inspector", v: insp.inspector_name || insp.inspector_username || "—" },
      { k: "Machine", v: `${insp.asset_code}${insp.asset_name ? ` — ${insp.asset_name}` : ""}` },
      { k: "Hour meter", v: insp.machine_hours != null ? Number(insp.machine_hours).toFixed(1) : "—" },
      { k: "Location", v: insp.location || "—" },
      { k: "Overall result", v: insp.overall_result ? label(OVERALL_RESULTS, insp.overall_result) : "—" },
      { k: "Work order", v: insp.work_order_id ? `#${insp.work_order_id}` : "—" },
    ], 2);

    const status = (c) => c.status || (c.ok === true ? "ok" : c.ok === false ? "fault" : "na");
    const faults = insp.checklist.filter((c) => status(c) === "fault");
    if (faults.length) {
      sectionTitle(doc, `Faults (${faults.length})`);
      for (const c of faults) {
        ensurePageSpace(doc, 30);
        doc.font("Helvetica-Bold").fontSize(10).fillColor("#b91c1c").text(`• ${c.label}`, left, doc.y, { width, continued: Boolean(c.note) });
        if (c.note) doc.font("Helvetica").fillColor("#111111").text(` — ${c.note}`);
        doc.moveDown(0.15);
      }
    }

    const groups = ARTISAN_SECTIONS.map((s) => ({ title: s.en, rows: insp.checklist.filter((c) => c.section === s.id) }));
    const loose = insp.checklist.filter((c) => !c.section || !ARTISAN_SECTIONS.some((s) => s.id === c.section));
    if (loose.length) groups.push({ title: "Checklist", rows: loose });
    for (const g of groups.filter((x) => x.rows.length)) {
      ensurePageSpace(doc, 70);
      sectionTitle(doc, g.title);
      for (const c of g.rows) {
        ensurePageSpace(doc, 40);
        const st = status(c);
        const y = doc.y;
        doc.font("Helvetica").fontSize(9.5).fillColor("#111111").text(c.label, left, y, { width: width - 70 });
        const after = doc.y;
        doc.font("Helvetica-Bold").fillColor(st === "fault" ? "#b91c1c" : st === "ok" ? "#15803d" : "#64748b")
          .text(STATUS_TEXT[st] || "—", left + width - 60, y, { width: 60, align: "right" });
        doc.y = Math.max(after, y + 12);
        if (c.note && st !== "fault") doc.font("Helvetica").fontSize(8.5).fillColor("#475569").text(c.note, left + 12, doc.y, { width: width - 80 });
        doc.moveDown(0.1);
      }
    }

    if (insp.notes) {
      sectionTitle(doc, "Notes");
      doc.font("Helvetica").fontSize(10).fillColor("#111111").text(insp.notes, left, doc.y, { width });
    }

    if (insp.photos?.length) {
      ensurePageSpace(doc, 260);
      sectionTitle(doc, `Photos (${insp.photos.length})`);
      for (const p of insp.photos) {
        const abs = resolveStorageAbs(p.file_path);
        ensurePageSpace(doc, 220);
        doc.font("Helvetica-Bold").fontSize(9.5).fillColor("#111111").text(p.caption || "Photo", left, doc.y, { width });
        doc.moveDown(0.2);
        const y = doc.y;
        if (abs && fs.existsSync(abs)) {
          try {
            const dims = doc.image(abs, left, y, { fit: [360, 190], align: "left", valign: "top" });
            doc.y = y + (dims?.height || 190) + 8;
          } catch {
            doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo could not be shown.");
          }
        } else {
          doc.font("Helvetica").fontSize(9).fillColor("#b91c1c").text("Photo file missing.");
        }
      }
    }
  }, {
    layout: "portrait",
    title: "IRONLOG",
    subtitle: "Artisan Inspection Report",
    rightText: `Inspection #${insp.id}`,
    showPageNumbers: true,
  });
}

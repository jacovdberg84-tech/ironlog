// IRONLOG/api/routes/safety.routes.js — safety equipment register, QR, inspections
import { db } from "../db/client.js";
import {
  DEFAULT_SAFETY_TEMPLATES,
  normalizeTemplateItems,
  parseTemplateItemsJson,
  buildChecklistFromTemplate,
  checklistStatus,
} from "../utils/safetyChecklistTemplates.js";
import { buildSafetyRegisterPdf } from "../utils/safetyRegisterPdf.js";

function isDate(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || "").trim());
}

function normalizeItemCode(raw) {
  return String(raw || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "-")
    .replace(/[^A-Z0-9_-]/g, "");
}

function isValidItemCode(code) {
  return /^[A-Z0-9][A-Z0-9_-]{2,31}$/.test(String(code || ""));
}

function resolveWebOrigin(req) {
  const proto = String(req.headers["x-forwarded-proto"] || req.protocol || "http").split(",")[0].trim();
  const host = String(req.headers["x-forwarded-host"] || req.headers.host || "").split(",")[0].trim();
  if (!host) return "";
  return `${proto}://${host}`;
}

function ensureSafetySchema() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS safety_checklist_templates (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      template_key TEXT NOT NULL,
      title TEXT NOT NULL,
      items_json TEXT NOT NULL DEFAULT '[]',
      site_code TEXT NOT NULL DEFAULT 'main',
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(template_key, site_code)
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS safety_equipment_items (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      item_code TEXT NOT NULL UNIQUE,
      template_key TEXT NOT NULL,
      item_name TEXT,
      location TEXT,
      notes TEXT,
      site_code TEXT NOT NULL DEFAULT 'main',
      active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS safety_inspections (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      item_id INTEGER NOT NULL,
      inspection_date TEXT NOT NULL,
      inspector_name TEXT,
      checklist_json TEXT NOT NULL,
      status TEXT NOT NULL DEFAULT 'pending',
      notes TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      UNIQUE(item_id, inspection_date),
      FOREIGN KEY (item_id) REFERENCES safety_equipment_items(id) ON DELETE CASCADE
    )
  `).run();

  db.prepare(`
    CREATE TABLE IF NOT EXISTS safety_equipment_qr_profiles (
      item_id INTEGER PRIMARY KEY,
      qr_payload TEXT NOT NULL,
      qr_text TEXT NOT NULL,
      generated_at TEXT NOT NULL,
      FOREIGN KEY (item_id) REFERENCES safety_equipment_items(id) ON DELETE CASCADE
    )
  `).run();

  const site_code = "main";
  const upsertTpl = db.prepare(`
    INSERT INTO safety_checklist_templates (template_key, title, items_json, site_code, active, created_at, updated_at)
    VALUES (?, ?, ?, ?, 1, datetime('now'), datetime('now'))
    ON CONFLICT(template_key, site_code) DO NOTHING
  `);
  for (const tpl of Object.values(DEFAULT_SAFETY_TEMPLATES)) {
    upsertTpl.run(
      tpl.template_key,
      tpl.title,
      JSON.stringify(normalizeTemplateItems(tpl.items)),
      site_code
    );
  }
}

function getTemplate(templateKey, siteCode = "main") {
  const key = String(templateKey || "").trim();
  const site = String(siteCode || "main").trim() || "main";
  const row = db.prepare(`
    SELECT id, template_key, title, items_json, site_code, active
    FROM safety_checklist_templates
    WHERE template_key = ? AND site_code = ? AND COALESCE(active, 1) = 1
  `).get(key, site);
  if (!row) return null;
  const items = parseTemplateItemsJson(row.items_json);
  return {
    id: Number(row.id),
    template_key: String(row.template_key),
    title: String(row.title),
    site_code: String(row.site_code),
    items,
  };
}

function getItemByCode(itemCode) {
  const code = normalizeItemCode(itemCode);
  if (!code) return null;
  return db.prepare(`
    SELECT id, item_code, template_key, item_name, location, notes, site_code, active
    FROM safety_equipment_items
    WHERE item_code = ? AND COALESCE(active, 1) = 1
  `).get(code);
}

function buildSafetyQrProfile(req, item) {
  const template = getTemplate(item.template_key, item.site_code);
  const origin = resolveWebOrigin(req);
  const path = `/web/safety-inspection.html?item_code=${encodeURIComponent(String(item.item_code))}`;
  const scan_url = origin ? `${origin}${path}` : path;
  const profile = {
    generated_at: new Date().toISOString(),
    item: {
      item_code: String(item.item_code),
      item_name: String(item.item_name || ""),
      location: String(item.location || ""),
      template_key: String(item.template_key),
      template_title: String(template?.title || item.template_key),
    },
    scan_url,
  };
  const qrText = [
    `IRONLOG SAFETY ${item.item_code}`,
    String(item.item_name || "").trim(),
    String(item.location || "").trim() ? `Location: ${item.location}` : "",
    `Scan URL: ${scan_url}`,
  ].filter(Boolean).join("\n");
  return { profile, qrText };
}

function storeQrProfile(itemId, profile, qrText) {
  db.prepare(`
    INSERT INTO safety_equipment_qr_profiles (item_id, qr_payload, qr_text, generated_at)
    VALUES (?, ?, ?, datetime('now'))
    ON CONFLICT(item_id) DO UPDATE SET
      qr_payload = excluded.qr_payload,
      qr_text = excluded.qr_text,
      generated_at = datetime('now')
  `).run(Number(itemId), JSON.stringify(profile), String(qrText || ""));
}

function latestInspectionForDate(itemId, date) {
  return db.prepare(`
    SELECT id, item_id, inspection_date, inspector_name, checklist_json, status, notes, created_at
    FROM safety_inspections
    WHERE item_id = ? AND inspection_date = ?
  `).get(Number(itemId), String(date));
}

function loadSafetyRegisterGroups(templateKey, checkDate) {
  const keyFilter = String(templateKey || "").trim();
  const templateRows = keyFilter
    ? db.prepare(`
        SELECT id, template_key, title, items_json, site_code
        FROM safety_checklist_templates
        WHERE template_key = ? AND COALESCE(active, 1) = 1
      `).all(keyFilter)
    : db.prepare(`
        SELECT id, template_key, title, items_json, site_code
        FROM safety_checklist_templates
        WHERE COALESCE(active, 1) = 1
        ORDER BY title ASC
      `).all();

  const groups = [];
  for (const row of templateRows) {
    const template = {
      template_key: String(row.template_key),
      title: String(row.title),
      site_code: String(row.site_code || "main"),
      items: parseTemplateItemsJson(row.items_json),
    };
    const itemRows = db.prepare(`
      SELECT id, item_code, template_key, item_name, location, notes, site_code
      FROM safety_equipment_items
      WHERE template_key = ? AND site_code = ? AND COALESCE(active, 1) = 1
      ORDER BY item_code ASC
    `).all(template.template_key, template.site_code);

    const items = itemRows.map((item) => {
      const insp = latestInspectionForDate(Number(item.id), checkDate);
      let checklist = [];
      if (insp?.checklist_json) {
        try {
          checklist = buildChecklistFromTemplate(template.items, JSON.parse(String(insp.checklist_json)));
        } catch {
          checklist = [];
        }
      }
      return { item, inspection: insp || null, checklist };
    });

    if (items.length) groups.push({ template, items });
  }
  return groups;
}

export default async function safetyRoutes(app) {
  ensureSafetySchema();

  app.get("/templates", async (_req, reply) => {
    try {
      const rows = db.prepare(`
        SELECT id, template_key, title, items_json, site_code, active, updated_at
        FROM safety_checklist_templates
        WHERE COALESCE(active, 1) = 1
        ORDER BY title ASC
      `).all();
      const templates = rows.map((r) => ({
        id: Number(r.id),
        template_key: String(r.template_key),
        title: String(r.title),
        site_code: String(r.site_code),
        items: parseTemplateItemsJson(r.items_json),
        updated_at: String(r.updated_at || ""),
      }));
      return reply.send({ ok: true, templates });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/templates/:template_key", async (req, reply) => {
    try {
      const template_key = String(req.params?.template_key || "").trim();
      const tpl = getTemplate(template_key);
      if (!tpl) return reply.code(404).send({ ok: false, error: "Template not found" });
      return reply.send({ ok: true, template: tpl });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/templates/:template_key", async (req, reply) => {
    try {
      const template_key = String(req.params?.template_key || "").trim();
      const existing = getTemplate(template_key);
      if (!existing) return reply.code(404).send({ ok: false, error: "Template not found" });
      const title = req.body?.title != null ? String(req.body.title || "").trim() : existing.title;
      const items = req.body?.items != null
        ? normalizeTemplateItems(req.body.items)
        : existing.items;
      if (!title) return reply.code(400).send({ ok: false, error: "title is required" });
      if (!items.length) return reply.code(400).send({ ok: false, error: "At least one checklist item is required" });
      db.prepare(`
        UPDATE safety_checklist_templates
        SET title = ?, items_json = ?, updated_at = datetime('now')
        WHERE template_key = ? AND site_code = ?
      `).run(title, JSON.stringify(items), template_key, existing.site_code);
      return reply.send({ ok: true, template: getTemplate(template_key, existing.site_code) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/items", async (_req, reply) => {
    try {
      const rows = db.prepare(`
        SELECT
          i.id,
          i.item_code,
          i.template_key,
          i.item_name,
          i.location,
          i.notes,
          i.site_code,
          i.active,
          i.created_at,
          i.updated_at,
          t.title AS template_title
        FROM safety_equipment_items i
        LEFT JOIN safety_checklist_templates t
          ON t.template_key = i.template_key AND t.site_code = i.site_code
        WHERE COALESCE(i.active, 1) = 1
        ORDER BY i.item_code ASC
      `).all();
      return reply.send({ ok: true, items: rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/items", async (req, reply) => {
    try {
      const item_code = normalizeItemCode(req.body?.item_code);
      const template_key = String(req.body?.template_key || "fire_extinguisher").trim();
      const item_name = String(req.body?.item_name || "").trim();
      const location = String(req.body?.location || "").trim();
      const notes = String(req.body?.notes || "").trim();
      if (!isValidItemCode(item_code)) {
        return reply.code(400).send({ ok: false, error: "item_code must be 3–32 chars (letters, numbers, - _)" });
      }
      if (!getTemplate(template_key)) {
        return reply.code(400).send({ ok: false, error: "Unknown template_key" });
      }
      const existing = db.prepare(`SELECT id FROM safety_equipment_items WHERE item_code = ?`).get(item_code);
      if (existing) return reply.code(409).send({ ok: false, error: "Item code already exists" });
      const ins = db.prepare(`
        INSERT INTO safety_equipment_items (
          item_code, template_key, item_name, location, notes, site_code, active, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, 'main', 1, datetime('now'), datetime('now'))
      `).run(item_code, template_key, item_name || null, location || null, notes || null);
      const row = db.prepare(`SELECT * FROM safety_equipment_items WHERE id = ?`).get(Number(ins.lastInsertRowid));
      return reply.send({ ok: true, item: row });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.put("/items/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const row = db.prepare(`SELECT * FROM safety_equipment_items WHERE id = ? AND COALESCE(active, 1) = 1`).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Item not found" });
      const template_key = req.body?.template_key != null
        ? String(req.body.template_key || "").trim()
        : String(row.template_key);
      if (!getTemplate(template_key)) {
        return reply.code(400).send({ ok: false, error: "Unknown template_key" });
      }
      const item_name = req.body?.item_name != null ? String(req.body.item_name || "").trim() : String(row.item_name || "");
      const location = req.body?.location != null ? String(req.body.location || "").trim() : String(row.location || "");
      const notes = req.body?.notes != null ? String(req.body.notes || "").trim() : String(row.notes || "");
      db.prepare(`
        UPDATE safety_equipment_items
        SET template_key = ?, item_name = ?, location = ?, notes = ?, updated_at = datetime('now')
        WHERE id = ?
      `).run(template_key, item_name || null, location || null, notes || null, id);
      const updated = db.prepare(`SELECT * FROM safety_equipment_items WHERE id = ?`).get(id);
      return reply.send({ ok: true, item: updated });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.delete("/items/:id", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });
      const row = db.prepare(`SELECT id, item_code FROM safety_equipment_items WHERE id = ?`).get(id);
      if (!row) return reply.code(404).send({ ok: false, error: "Item not found" });
      db.prepare(`UPDATE safety_equipment_items SET active = 0, updated_at = datetime('now') WHERE id = ?`).run(id);
      return reply.send({ ok: true, id, item_code: String(row.item_code) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/hub", async (req, reply) => {
    try {
      const check_date = String(req.query?.date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!isDate(check_date)) return reply.code(400).send({ ok: false, error: "date must be YYYY-MM-DD" });
      const items = db.prepare(`
        SELECT id, item_code, template_key, item_name, location, site_code
        FROM safety_equipment_items
        WHERE COALESCE(active, 1) = 1
        ORDER BY item_code ASC
      `).all();
      const out = [];
      let compliant = 0;
      for (const item of items) {
        const tpl = getTemplate(item.template_key, item.site_code);
        const insp = latestInspectionForDate(Number(item.id), check_date);
        const status = insp ? String(insp.status || "pending") : "pending";
        const done = Boolean(insp && ["pass", "attention"].includes(status));
        if (done) compliant += 1;
        out.push({
          item_id: Number(item.id),
          item_code: String(item.item_code),
          item_name: String(item.item_name || ""),
          location: String(item.location || ""),
          template_key: String(item.template_key),
          template_title: String(tpl?.title || item.template_key),
          status: done ? "compliant" : "pending",
          inspection_status: status,
          inspection_id: insp ? Number(insp.id) : null,
        });
      }
      return reply.send({
        ok: true,
        check_date,
        items: out,
        summary: {
          total: out.length,
          compliant,
          pending: Math.max(0, out.length - compliant),
        },
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/inspection-context", async (req, reply) => {
    try {
      const item_code = normalizeItemCode(req.query?.item_code);
      const inspection_date = String(req.query?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      if (!item_code) return reply.code(400).send({ ok: false, error: "item_code is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });
      const item = getItemByCode(item_code);
      if (!item) return reply.code(404).send({ ok: false, error: "Safety item not found" });
      const template = getTemplate(item.template_key, item.site_code);
      if (!template) return reply.code(404).send({ ok: false, error: "Checklist template not found" });
      const existing = latestInspectionForDate(Number(item.id), inspection_date);
      let checklist = buildChecklistFromTemplate(template.items, []);
      if (existing?.checklist_json) {
        try {
          checklist = buildChecklistFromTemplate(template.items, JSON.parse(String(existing.checklist_json)));
        } catch {
          checklist = buildChecklistFromTemplate(template.items, []);
        }
      }
      return reply.send({
        ok: true,
        inspection_date,
        item: {
          id: Number(item.id),
          item_code: String(item.item_code),
          item_name: String(item.item_name || ""),
          location: String(item.location || ""),
          template_key: String(item.template_key),
        },
        template,
        existing_inspection: existing
          ? {
              id: Number(existing.id),
              inspector_name: String(existing.inspector_name || ""),
              notes: String(existing.notes || ""),
              status: String(existing.status || ""),
              checklist,
            }
          : null,
        checklist,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/inspections", async (req, reply) => {
    try {
      const item_code = normalizeItemCode(req.body?.item_code);
      const inspection_date = String(req.body?.inspection_date || "").trim() || new Date().toISOString().slice(0, 10);
      const inspector_name = String(req.body?.inspector_name || "").trim();
      const notes = String(req.body?.notes || "").trim();
      if (!item_code) return reply.code(400).send({ ok: false, error: "item_code is required" });
      if (!isDate(inspection_date)) return reply.code(400).send({ ok: false, error: "inspection_date must be YYYY-MM-DD" });
      const item = getItemByCode(item_code);
      if (!item) return reply.code(404).send({ ok: false, error: "Safety item not found" });
      const template = getTemplate(item.template_key, item.site_code);
      if (!template) return reply.code(404).send({ ok: false, error: "Checklist template not found" });
      const checklist = buildChecklistFromTemplate(template.items, req.body?.checklist);
      const unanswered = checklist.filter((r) => r.ok == null);
      if (unanswered.length) {
        return reply.code(400).send({
          ok: false,
          error: `Complete all checklist items (${unanswered.length} unanswered)`,
        });
      }
      const failed = checklist.filter((r) => r.ok === false);
      for (const row of failed) {
        if (!String(row.note || "").trim()) {
          return reply.code(400).send({
            ok: false,
            error: `Add a note for failed item: ${row.label}`,
          });
        }
      }
      const status = checklistStatus(checklist);
      const checklist_json = JSON.stringify(checklist);
      db.prepare(`
        INSERT INTO safety_inspections (
          item_id, inspection_date, inspector_name, checklist_json, status, notes, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, datetime('now'), datetime('now'))
        ON CONFLICT(item_id, inspection_date) DO UPDATE SET
          inspector_name = excluded.inspector_name,
          checklist_json = excluded.checklist_json,
          status = excluded.status,
          notes = excluded.notes,
          updated_at = datetime('now')
      `).run(
        Number(item.id),
        inspection_date,
        inspector_name || null,
        checklist_json,
        status,
        notes || null
      );
      const saved = latestInspectionForDate(Number(item.id), inspection_date);
      return reply.send({ ok: true, inspection: saved, checklist, status });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/items/:item_code/qr-profile", async (req, reply) => {
    try {
      const item = getItemByCode(req.params?.item_code);
      if (!item) return reply.code(404).send({ ok: false, error: "Safety item not found" });
      const stored = db.prepare(`
        SELECT qr_payload, qr_text, generated_at FROM safety_equipment_qr_profiles WHERE item_id = ?
      `).get(Number(item.id));
      if (stored?.qr_payload) {
        try {
          const payload = JSON.parse(String(stored.qr_payload));
          return reply.send({
            ok: true,
            qr_payload: payload,
            qr_text: String(stored.qr_text || ""),
            generated_at: String(stored.generated_at || ""),
            cached: true,
          });
        } catch {
          /* refresh below */
        }
      }
      const built = buildSafetyQrProfile(req, item);
      storeQrProfile(Number(item.id), built.profile, built.qrText);
      return reply.send({
        ok: true,
        qr_payload: built.profile,
        qr_text: built.qrText,
        generated_at: built.profile.generated_at,
        cached: false,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/items/:item_code/qr-profile/refresh", async (req, reply) => {
    try {
      const item = getItemByCode(req.params?.item_code);
      if (!item) return reply.code(404).send({ ok: false, error: "Safety item not found" });
      const built = buildSafetyQrProfile(req, item);
      storeQrProfile(Number(item.id), built.profile, built.qrText);
      return reply.send({
        ok: true,
        qr_payload: built.profile,
        qr_text: built.qrText,
        generated_at: built.profile.generated_at,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  // GET /api/safety/register.pdf?template_key=fire_extinguisher&date=YYYY-MM-DD&blank=1&download=1
  app.get("/register.pdf", async (req, reply) => {
    try {
      const template_key = String(req.query?.template_key || "").trim();
      const check_date = String(req.query?.date || "").trim() || new Date().toISOString().slice(0, 10);
      const blank = String(req.query?.blank || "").trim() === "1";
      const download = String(req.query?.download || "").trim() === "1";
      if (!isDate(check_date)) {
        return reply.code(400).send({ ok: false, error: "date must be YYYY-MM-DD" });
      }
      if (template_key && !getTemplate(template_key)) {
        return reply.code(404).send({ ok: false, error: "Template not found" });
      }

      const groups = loadSafetyRegisterGroups(template_key, check_date);
      const pdf = await buildSafetyRegisterPdf({ groups, checkDate: check_date, blank });
      const slug = template_key ? template_key.replace(/[^a-z0-9_-]+/gi, "_") : "all_types";
      const mode = blank ? "blank" : "register";

      return reply
        .header("Content-Type", "application/pdf")
        .header(
          "Content-Disposition",
          `${download ? "attachment" : "inline"}; filename="AML_Safety_${mode}_${slug}_${check_date}.pdf"`
        )
        .send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
}

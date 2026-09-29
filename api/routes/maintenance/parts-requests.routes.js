// IRONLOG/api/routes/maintenance/parts-requests.routes.js — Workshop parts requests and RFQ PDF.
// Registered by routes/maintenance.routes.js; shared helpers arrive through ctx.
import { buildPdfBuffer, sectionTitle, table } from "../../utils/pdfGenerator.js";
import { db } from "../../db/client.js";
import { getPdfReportBranding } from "../../utils/reportSettings.js";
import { writeAudit } from "../../utils/audit.js";

export default function registerPartsRequestsRoutes(app, ctx) {
  const { PARTS_REQUEST_MANAGERS, PARTS_REQUEST_STATUSES, PARTS_REQUEST_URGENCY, getMaintenanceRoles } = ctx;

  // GET /api/maintenance/parts-requests/rfq.pdf?ids=1,2,3&supplier=...&reference=...
  app.get("/parts-requests/rfq.pdf", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const idsRaw = req.query?.ids;
      const requestIds = Array.isArray(idsRaw)
        ? [...new Set(idsRaw.map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))]
        : [...new Set(String(idsRaw || "").split(/[,\s]+/).map((v) => Number(v)).filter((v) => Number.isInteger(v) && v > 0))];
      if (!requestIds.length) {
        return reply.code(400).send({ ok: false, error: "Select at least one parts request (ids)" });
      }

      const supplier = String(req.query?.supplier || "").trim();
      if (!supplier) {
        return reply.code(400).send({ ok: false, error: "supplier is required" });
      }

      const reference = String(req.query?.reference || "").trim() || `RFQ-${new Date().toISOString().slice(0, 10)}`;
      const contact = String(req.query?.contact || "").trim();
      const email = String(req.query?.email || "").trim();
      const phone = String(req.query?.phone || "").trim();
      const required_by = String(req.query?.required_by || "").trim();
      const notes = String(req.query?.notes || "").trim();
      const requested_by = String(req.query?.requested_by || req.headers?.["x-user-name"] || "").trim() || "system";
      const asOfLabel = new Date().toISOString().slice(0, 10);
      const branding = getPdfReportBranding(db);

      const placeholders = requestIds.map(() => "?").join(", ");
      const rows = db.prepare(`
        SELECT
          pr.id,
          pr.asset_code,
          pr.part_code,
          pr.part_name,
          pr.qty,
          pr.urgency,
          pr.notes,
          pr.work_order_id,
          pr.status,
          pr.requested_by,
          pr.created_at,
          a.asset_name
        FROM maintenance_parts_requests pr
        LEFT JOIN assets a ON a.id = pr.asset_id
        WHERE pr.id IN (${placeholders})
          AND COALESCE(pr.site_code, 'main') = ?
        ORDER BY
          CASE LOWER(COALESCE(pr.urgency, 'normal'))
            WHEN 'critical' THEN 0
            WHEN 'urgent' THEN 1
            ELSE 2
          END ASC,
          pr.asset_code ASC,
          pr.part_code ASC,
          pr.id ASC
      `).all(...requestIds, site_code);

      if (!rows.length) {
        return reply.code(404).send({ ok: false, error: "No matching parts requests found" });
      }

      const urgencyLabel = (u) => {
        const x = String(u || "normal").toLowerCase();
        if (x === "critical") return "Critical";
        if (x === "urgent") return "Urgent";
        return "Normal";
      };

      const lineRows = rows.map((r, idx) => ({
        line_no: String(idx + 1),
        part_code: String(r.part_code || "—"),
        part_name: String(r.part_name || "—"),
        qty: Number(r.qty || 0).toFixed(1).replace(/\.0$/, ""),
        asset: [r.asset_code, r.asset_name].map((x) => String(x || "").trim()).filter(Boolean).join(" — ") || "—",
        work_order_id: r.work_order_id ? String(r.work_order_id) : "—",
        urgency: urgencyLabel(r.urgency),
        notes: String(r.notes || "").trim() || "—",
      }));

      const pdf = await buildPdfBuffer(
        (doc) => {
          sectionTitle(doc, "Request for Quote");
          doc.font("Helvetica").fontSize(10).fillColor("#0f172a");
          doc.text(`Reference: ${reference}`);
          doc.text(`Date: ${asOfLabel}`);
          if (branding.company_name) doc.text(`From: ${branding.company_name}${branding.site_name ? ` — ${branding.site_name}` : ""}`);
          doc.text(`Prepared by: ${requested_by}`);
          doc.moveDown(0.35);
          doc.font("Helvetica-Bold").text("To:");
          doc.font("Helvetica");
          doc.text(supplier);
          if (contact) doc.text(`Attention: ${contact}`);
          if (email) doc.text(`Email: ${email}`);
          if (phone) doc.text(`Phone: ${phone}`);
          if (required_by) doc.text(`Quote required by: ${required_by}`);
          doc.moveDown(0.5);
          doc.font("Helvetica").fontSize(10).text(
            "Please provide pricing, availability, and lead time for the parts listed below.",
          );
          doc.moveDown(0.4);

          table(
            doc,
            [
              { key: "line_no", label: "#", width: 0.04, align: "center" },
              { key: "part_code", label: "Part code", width: 0.12 },
              { key: "part_name", label: "Description", width: 0.24 },
              { key: "qty", label: "Qty", width: 0.06, align: "right" },
              { key: "asset", label: "Asset", width: 0.16 },
              { key: "work_order_id", label: "WO #", width: 0.07, align: "center" },
              { key: "urgency", label: "Urgency", width: 0.09 },
              { key: "notes", label: "Notes", width: 0.22 },
            ],
            lineRows,
          );

          if (notes) {
            doc.moveDown(0.6);
            sectionTitle(doc, "Additional notes / terms");
            doc.font("Helvetica").fontSize(10).text(notes, { width: doc.page.width - doc.page.margins.left - doc.page.margins.right });
          }

          doc.moveDown(1.2);
          doc.font("Helvetica").fontSize(10);
          doc.text("Authorized by: ________________________________     Date: ________________");
        },
        {
          title: branding.company_name || "IRONLOG",
          subtitle: "Parts Request for Quote",
          rightText: reference,
          layout: "landscape",
          db,
        },
      );

      const isDownload = String(req.query?.download || "").trim() === "1";
      const safeRef = reference.replace(/[^\w.-]+/g, "_");
      reply.header("Cache-Control", "no-store, no-cache, must-revalidate");
      reply.header("Pragma", "no-cache");
      reply.header("Content-Type", "application/pdf");
      reply.header(
        "Content-Disposition",
        `${isDownload ? "attachment" : "inline"}; filename="${safeRef}.pdf"`,
      );
      return reply.send(pdf);
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.get("/parts-requests", async (req, reply) => {
    try {
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const status = String(req.query?.status || "").trim().toLowerCase();
      const mine = String(req.query?.mine || "").trim() === "1";
      const userName = String(req.headers?.["x-user-name"] || "").trim();
      const limitRaw = Number(req.query?.limit || 500);
      const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(2000, Math.trunc(limitRaw))) : 500;

      const where = ["COALESCE(pr.site_code, 'main') = ?"];
      const params = [site_code];
      if (status && PARTS_REQUEST_STATUSES.has(status)) {
        where.push("LOWER(COALESCE(pr.status, 'requested')) = ?");
        params.push(status);
      }
      if (mine && userName) {
        where.push("LOWER(COALESCE(pr.requested_by, '')) = ?");
        params.push(userName.toLowerCase());
      }

      const rows = db.prepare(`
        SELECT
          pr.id,
          pr.site_code,
          pr.asset_id,
          pr.asset_code,
          pr.part_code,
          pr.part_name,
          pr.qty,
          pr.urgency,
          pr.notes,
          pr.work_order_id,
          pr.status,
          pr.requested_by,
          pr.ordered_by,
          pr.status_notes,
          pr.created_at,
          pr.updated_at,
          a.asset_name
        FROM maintenance_parts_requests pr
        LEFT JOIN assets a ON a.id = pr.asset_id
        WHERE ${where.join(" AND ")}
        ORDER BY
          CASE LOWER(COALESCE(pr.status, 'requested'))
            WHEN 'requested' THEN 0
            WHEN 'ordered' THEN 1
            WHEN 'received' THEN 2
            ELSE 3
          END ASC,
          CASE LOWER(COALESCE(pr.urgency, 'normal'))
            WHEN 'critical' THEN 0
            WHEN 'urgent' THEN 1
            ELSE 2
          END ASC,
          pr.created_at DESC,
          pr.id DESC
        LIMIT ${limit}
      `).all(...params);
      return reply.send({ ok: true, rows });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.post("/parts-requests", async (req, reply) => {
    try {
      const body = req.body || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const requested_by = String(req.headers?.["x-user-name"] || body.requested_by || "system").trim() || "system";
      const part_code = String(body.part_code || "").trim();
      let part_name = String(body.part_name || "").trim();
      const qty = Number(body.qty ?? 1);
      const urgencyRaw = String(body.urgency || "normal").trim().toLowerCase();
      const urgency = PARTS_REQUEST_URGENCY.has(urgencyRaw) ? urgencyRaw : "normal";
      const notes = String(body.notes || "").trim();
      const work_order_id = Number(body.work_order_id || 0) || null;
      const asset_id = Number(body.asset_id || 0) || null;

      if (!part_code && !part_name) {
        return reply.code(400).send({ ok: false, error: "Part description or part code is required" });
      }
      if (!Number.isFinite(qty) || qty <= 0) {
        return reply.code(400).send({ ok: false, error: "qty must be greater than zero" });
      }

      if (part_code && !part_name) {
        const hit = db.prepare(`SELECT part_name FROM parts WHERE LOWER(part_code) = LOWER(?) LIMIT 1`).get(part_code);
        if (hit?.part_name) part_name = String(hit.part_name).trim();
      }
      if (!part_name) part_name = part_code;

      let asset_code = String(body.asset_code || "").trim();
      if (asset_id) {
        const asset = db.prepare(`SELECT id, asset_code FROM assets WHERE id = ? LIMIT 1`).get(asset_id);
        if (!asset) return reply.code(400).send({ ok: false, error: "Invalid asset" });
        asset_code = String(asset.asset_code || asset_code).trim();
      }

      const now = new Date().toISOString();
      const info = db.prepare(`
        INSERT INTO maintenance_parts_requests (
          site_code, asset_id, asset_code, part_code, part_name, qty, urgency, notes,
          work_order_id, status, requested_by, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'requested', ?, ?, ?)
      `).run(
        site_code,
        asset_id,
        asset_code || null,
        part_code || null,
        part_name,
        qty,
        urgency,
        notes || null,
        work_order_id,
        requested_by,
        now,
        now
      );

      writeAudit(db, req, {
        module: "maintenance",
        action: "parts_request.create",
        entity_type: "maintenance_parts_request",
        entity_id: String(info.lastInsertRowid),
        after: { part_code, part_name, qty, urgency, asset_code, work_order_id },
      });

      return reply.send({ ok: true, id: Number(info.lastInsertRowid || 0) });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });

  app.patch("/parts-requests/:id/status", async (req, reply) => {
    try {
      const id = Number(req.params?.id || 0);
      if (!id) return reply.code(400).send({ ok: false, error: "Invalid id" });

      const body = req.body || {};
      const site_code = String(req.headers?.["x-site-code"] || "main").trim().toLowerCase() || "main";
      const userName = String(req.headers?.["x-user-name"] || "").trim() || "system";
      const roles = getMaintenanceRoles(req);
      const isManager = roles.some((r) => PARTS_REQUEST_MANAGERS.includes(r));

      const status = String(body.status || "").trim().toLowerCase();
      if (!PARTS_REQUEST_STATUSES.has(status)) {
        return reply.code(400).send({ ok: false, error: "Invalid status" });
      }

      const existing = db.prepare(`
        SELECT id, status, requested_by
        FROM maintenance_parts_requests
        WHERE id = ? AND COALESCE(site_code, 'main') = ?
      `).get(id, site_code);
      if (!existing) return reply.code(404).send({ ok: false, error: "Request not found" });

      if (!isManager) {
        const own = String(existing.requested_by || "").trim().toLowerCase() === userName.toLowerCase();
        if (!own || status !== "cancelled" || String(existing.status || "").toLowerCase() !== "requested") {
          return reply.code(403).send({ ok: false, error: "not allowed" });
        }
      }

      const status_notes = String(body.status_notes || "").trim() || null;
      const ordered_by = ["ordered", "received"].includes(status) ? userName : null;
      const now = new Date().toISOString();

      db.prepare(`
        UPDATE maintenance_parts_requests
        SET status = ?, status_notes = ?, ordered_by = COALESCE(?, ordered_by), updated_at = ?
        WHERE id = ? AND COALESCE(site_code, 'main') = ?
      `).run(status, status_notes, ordered_by, now, id, site_code);

      writeAudit(db, req, {
        module: "maintenance",
        action: "parts_request.status",
        entity_type: "maintenance_parts_request",
        entity_id: String(id),
        after: { status, status_notes },
      });

      return reply.send({ ok: true, id });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message || String(err) });
    }
  });
}

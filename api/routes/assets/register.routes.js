// IRONLOG/api/routes/assets/register.routes.js — Asset register, fleet summary, hire register, cost centre export, create, archive and edit.
// Registered by routes/assets.routes.js; shared helpers arrive through ctx.
import ExcelJS from "exceljs";
import { db } from "../../db/client.js";
import { getAssetCurrentHoursInfo } from "../../utils/assetMeterHours.js";
import { listHireAssetsRegister, normalizeHireBillingMode } from "../../utils/plantHire.js";
import { normalizeCostCenterCode, normalizeSiteCode } from "../../utils/costAllocation.js";
import { validateAgainstMdmPolicy, validateAssetGovernanceOptional } from "../../utils/masterdataGovernance.js";

export default function registerRegisterRoutes(app, ctx) {
  const { buildMachineStatus, getAssetByCode, hasTable, insertAsset, siteCodeFromReq } = ctx;

  /* =========================
     ROUTES
  ========================= */

  // GET /api/assets?include_archived=1
  app.get("/", async (req) => {
    const includeArchived = String(req.query?.include_archived || "0") === "1";

    const rows = db.prepare(`
      SELECT
        id, asset_code, asset_name, category,
        active, is_standby,
        archived, archive_reason, archived_at,
        department_code, cost_center_code, site_code, data_owner_username,
        created_at
      FROM assets
      WHERE (? = 1 OR archived = 0)
      ORDER BY asset_code ASC
    `).all(includeArchived ? 1 : 0);

    return rows.map((r) => ({
      ...r,
      active: Number(r.active),
      is_standby: Number(r.is_standby),
      archived: Number(r.archived),
      site_code: r.site_code != null ? String(r.site_code) : null,
      cost_center_code: r.cost_center_code != null ? String(r.cost_center_code) : null,
      department_code: r.department_code != null ? String(r.department_code) : null,
    }));
  });

  // GET /api/assets/fleet-summary?include_archived=1
  app.get("/fleet-summary", async (req) => {
    const includeArchived = String(req.query?.include_archived || "0") === "1";

    const rows = db.prepare(`
      SELECT
        id, asset_code, asset_name, category,
        active, is_standby,
        archived, archive_reason
      FROM assets
      WHERE (? = 1 OR archived = 0)
      ORDER BY asset_code ASC
    `).all(includeArchived ? 1 : 0);

    const fuel30Stmt = hasTable("fuel_logs")
      ? db.prepare(`
          SELECT COALESCE(SUM(liters), 0) AS liters_30d
          FROM fuel_logs
          WHERE asset_id = ?
            AND log_date >= date('now', '-30 days')
        `)
      : null;

    const cards = rows.map((a) => {
      const meter = getAssetCurrentHoursInfo(a.id);
      const fuelRow = fuel30Stmt ? fuel30Stmt.get(a.id) : null;
      return {
        asset_code: a.asset_code,
        asset_name: a.asset_name,
        category: a.category,
        active: Number(a.active),
        is_standby: Number(a.is_standby),
        archived: Number(a.archived),
        archive_reason: a.archive_reason || null,
        current_hours: Number(Number(meter.hours || 0).toFixed(1)),
        status: buildMachineStatus(a.id),
        fuel_liters_30d: Number(Number(fuelRow?.liters_30d || 0).toFixed(1)),
      };
    });

    return { ok: true, cards };
  });

  // GET /api/assets/hire-register
  app.get("/hire-register", async () => {
    const rows = listHireAssetsRegister(db).map((r) => ({
      asset_code: r.asset_code,
      asset_name: r.asset_name,
      category: r.category,
      site_code: r.site_code,
      cost_center_code: r.cost_center_code,
      hire_billing_mode: r.hire_billing_mode || "",
      hire_rate_per_hour: r.hire_rate_per_hour != null ? Number(r.hire_rate_per_hour) : null,
      hire_fixed_monthly: r.hire_fixed_monthly != null ? Number(r.hire_fixed_monthly) : null,
      utilization_mode: r.utilization_mode || "",
      active: Number(r.active),
      archived: Number(r.archived),
    }));
    return { ok: true, rows };
  });

  // GET /api/assets/cost-centers.xlsx?include_archived=0|1
  // Equipment register for manual cost center allocation in Excel.
  app.get("/cost-centers.xlsx", async (req, reply) => {
    const includeArchived = String(req.query?.include_archived || "0") === "1";
    const today = new Date().toISOString().slice(0, 10);

    const rows = db.prepare(`
      SELECT
        asset_code,
        asset_name,
        category,
        site_code,
        department_code,
        cost_center_code,
        active,
        is_standby,
        archived,
        archive_reason
      FROM assets
      WHERE (? = 1 OR COALESCE(archived, 0) = 0)
      ORDER BY asset_code ASC
    `).all(includeArchived ? 1 : 0);

    const wb = new ExcelJS.Workbook();
    wb.creator = "IRONLOG";
    wb.created = new Date();

    const wsInfo = wb.addWorksheet("Instructions");
    wsInfo.columns = [
      { header: "Topic", key: "topic", width: 22 },
      { header: "Detail", key: "detail", width: 72 },
    ];
    wsInfo.getRow(1).font = { bold: true };
    wsInfo.addRows([
      {
        topic: "Purpose",
        detail: "Fill in the Cost Center Code column for each asset. Other columns are for reference.",
      },
      {
        topic: "Asset Code",
        detail: "Do not change asset codes — they are the key used when importing updates later.",
      },
      {
        topic: "Cost Center Code",
        detail: "Use codes from Enterprise → Master Data → Cost Centers (e.g. MINE-OPS-01).",
      },
      {
        topic: "Apply changes",
        detail: "For now, copy completed codes into IRONLOG via Assets → Site & Cost Center, or ask admin to import.",
      },
      {
        topic: "Exported",
        detail: `${today} | ${rows.length} asset(s) | Archived included: ${includeArchived ? "yes" : "no"}`,
      },
    ]);

    const ws = wb.addWorksheet("Equipment");
    ws.columns = [
      { header: "Asset Code", key: "asset_code", width: 14 },
      { header: "Asset Name", key: "asset_name", width: 28 },
      { header: "Category", key: "category", width: 22 },
      { header: "Site Code", key: "site_code", width: 12 },
      { header: "Department Code", key: "department_code", width: 16 },
      { header: "Cost Center Code", key: "cost_center_code", width: 18 },
      { header: "Active", key: "active", width: 10 },
      { header: "Standby", key: "is_standby", width: 10 },
      { header: "Archived", key: "archived", width: 10 },
      { header: "Archive Reason", key: "archive_reason", width: 24 },
    ];
    ws.getRow(1).font = { bold: true };
    ws.views = [{ state: "frozen", ySplit: 1 }];

    for (const r of rows) {
      ws.addRow({
        asset_code: r.asset_code || "",
        asset_name: r.asset_name || "",
        category: r.category || "",
        site_code: r.site_code != null ? String(r.site_code) : "",
        department_code: r.department_code != null ? String(r.department_code) : "",
        cost_center_code: r.cost_center_code != null ? String(r.cost_center_code) : "",
        active: Number(r.active) ? "Yes" : "No",
        is_standby: Number(r.is_standby) ? "Yes" : "No",
        archived: Number(r.archived) ? "Yes" : "No",
        archive_reason: r.archive_reason || "",
      });
    }

    const ccCol = ws.getColumn("cost_center_code");
    ccCol.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
      if (rowNumber === 1) return;
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFF9E6" },
      };
    });

    const buffer = await wb.xlsx.writeBuffer();
    reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
      .header(
        "Content-Disposition",
        `attachment; filename="IRONLOG_Asset_Cost_Centers_${today}.xlsx"`,
      )
      .send(Buffer.from(buffer));
  });

  // POST /api/assets
  app.post("/", async (req, reply) => {
    const body = req.body || {};
    const asset_code = String(body.asset_code || "").trim();
    const asset_name = String(body.asset_name || "").trim();
    const category = String(body.category || "").trim() || null;

    const active = body.active === 0 || body.active === false ? 0 : 1;
    const is_standby = body.is_standby ? 1 : 0;

    // archived fields (optional)
    const archived = body.archived ? 1 : 0;
    const archive_reason = archived ? (String(body.archive_reason || "").trim() || null) : null;
    const archived_at = archived ? (String(body.archived_at || "").trim() || null) : null;

    const department_code =
      body.department_code != null && String(body.department_code).trim() !== ""
        ? String(body.department_code).trim().toUpperCase()
        : null;
    const cost_center_code = normalizeCostCenterCode(body.cost_center_code);
    const site_code = normalizeSiteCode(body.site_code) || siteCodeFromReq(req);
    const data_owner_username =
      body.data_owner_username != null && String(body.data_owner_username).trim() !== ""
        ? String(body.data_owner_username).trim()
        : null;

    if (!asset_code || !asset_name) {
      return reply.code(400).send({ error: "asset_code and asset_name are required" });
    }

    const policyBody = {
      department_code,
      cost_center_code,
      data_owner_username,
    };
    const pol = validateAgainstMdmPolicy(siteCodeFromReq(req), "asset", policyBody);
    if (!pol.ok) {
      return reply.code(400).send({ error: `missing required fields: ${pol.missing.join(", ")}` });
    }

    const gov = validateAssetGovernanceOptional(siteCodeFromReq(req), {
      department_code,
      cost_center_code,
    });
    if (!gov.ok) return reply.code(400).send({ error: gov.error });

    try {
      const r = insertAsset.run(
        asset_code,
        asset_name,
        category,
        active,
        is_standby,
        archived,
        archive_reason,
        archived_at,
        department_code,
        cost_center_code,
        site_code,
        data_owner_username
      );
      return reply.code(201).send({ ok: true, id: Number(r.lastInsertRowid) });
    } catch (e) {
      return reply.code(400).send({ error: e.message || String(e) });
    }
  });

  // POST /api/assets/:asset_code/archive  { archived: true/false, reason? }
  app.post("/:asset_code/archive", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    const body = req.body || {};

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const archived = body.archived ? 1 : 0;
    const reason = archived ? (String(body.reason || "").trim() || null) : null;

    if (archived) {
      db.prepare(`
        UPDATE assets
        SET archived = 1, archive_reason = ?, archived_at = datetime('now')
        WHERE asset_code = ?
      `).run(reason, asset_code);
    } else {
      db.prepare(`
        UPDATE assets
        SET archived = 0, archive_reason = NULL, archived_at = NULL
        WHERE asset_code = ?
      `).run(asset_code);
    }

    return { ok: true };
  });

  // PATCH /api/assets/:asset_code  { active?, is_standby?, site_code?, cost_center_code?, department_code?, category?, asset_name? }
  app.patch("/:asset_code", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    const body = req.body || {};

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const sets = [];
    const args = [];
    if (body.active !== undefined) {
      sets.push("active = ?");
      args.push(body.active ? 1 : 0);
    }
    if (body.is_standby !== undefined) {
      sets.push("is_standby = ?");
      args.push(body.is_standby ? 1 : 0);
    }
    if (body.site_code !== undefined) {
      sets.push("site_code = ?");
      args.push(normalizeSiteCode(body.site_code));
    }
    if (body.cost_center_code !== undefined) {
      sets.push("cost_center_code = ?");
      args.push(normalizeCostCenterCode(body.cost_center_code));
    }
    if (body.department_code !== undefined) {
      const dept = String(body.department_code || "").trim();
      sets.push("department_code = ?");
      args.push(dept ? dept.toUpperCase() : null);
    }
    if (body.category !== undefined) {
      const cat = String(body.category || "").trim();
      sets.push("category = ?");
      args.push(cat || null);
    }
    if (body.asset_name !== undefined) {
      const name = String(body.asset_name || "").trim();
      if (!name) return reply.code(400).send({ error: "asset_name cannot be empty" });
      sets.push("asset_name = ?");
      args.push(name);
    }
    if (body.hire_billing_mode !== undefined) {
      sets.push("hire_billing_mode = ?");
      args.push(normalizeHireBillingMode(body.hire_billing_mode) || null);
    }
    if (body.hire_rate_per_hour !== undefined) {
      const v = body.hire_rate_per_hour === null || String(body.hire_rate_per_hour).trim() === ""
        ? null
        : Number(body.hire_rate_per_hour);
      if (v != null && (!Number.isFinite(v) || v < 0)) {
        return reply.code(400).send({ error: "hire_rate_per_hour must be >= 0" });
      }
      sets.push("hire_rate_per_hour = ?");
      args.push(v);
    }
    if (body.hire_fixed_monthly !== undefined) {
      const v = body.hire_fixed_monthly === null || String(body.hire_fixed_monthly).trim() === ""
        ? null
        : Number(body.hire_fixed_monthly);
      if (v != null && (!Number.isFinite(v) || v < 0)) {
        return reply.code(400).send({ error: "hire_fixed_monthly must be >= 0" });
      }
      sets.push("hire_fixed_monthly = ?");
      args.push(v);
    }

    if (body.cost_center_code !== undefined || body.department_code !== undefined) {
      const gov = validateAssetGovernanceOptional(siteCodeFromReq(req), {
        department_code: body.department_code !== undefined
          ? (String(body.department_code || "").trim() ? String(body.department_code).trim().toUpperCase() : null)
          : undefined,
        cost_center_code: body.cost_center_code !== undefined
          ? normalizeCostCenterCode(body.cost_center_code)
          : undefined,
      });
      if (!gov.ok) return reply.code(400).send({ error: gov.error });
    }

    const tx = db.transaction(() => {
      if (sets.length) {
        db.prepare(`UPDATE assets SET ${sets.join(", ")} WHERE asset_code = ?`).run(...args, asset_code);
      }
    });

    tx();

    const updated = db.prepare(`
      SELECT id, asset_code, asset_name, category, active, is_standby, archived,
             department_code, cost_center_code, site_code, data_owner_username, created_at,
             hire_billing_mode, hire_rate_per_hour, hire_fixed_monthly
      FROM assets WHERE asset_code = ?
    `).get(asset_code);
    return {
      ok: true,
      asset: {
        ...updated,
        active: Number(updated.active),
        is_standby: Number(updated.is_standby),
        archived: Number(updated.archived),
      },
    };
  });
}

// IRONLOG/api/routes/assets/history.routes.js — Asset hours and history.
// Registered by routes/assets.routes.js; shared helpers arrive through ctx.
import { getAssetCurrentHoursInfo } from "../../utils/assetMeterHours.js";
import { db } from "../../db/client.js";
import { assetHistoryWorkbook, buildAssetHistory } from "../../utils/assetHistory.js";
import { isDate } from "../../utils/request.js";
import { andSite } from "../../utils/sites.js";

export default function registerHistoryRoutes(app, ctx) {
  const { getAssetByCode, getAssetById } = ctx;

  // ------------------------------------------------------------
  // Hour meter snapshot for one machine
  // GET /api/assets/:asset_id_or_asset_code/hours
  // Resolves numeric segment as asset id first, then asset_code.
  // ------------------------------------------------------------
  app.get("/:identifier/hours", async (req, reply) => {
    const identifier = String(req.params.identifier || "").trim();
    if (!identifier) return reply.code(404).send({ error: "Asset not found" });

    let asset = null;
    if (/^\d+$/.test(identifier)) {
      asset = getAssetById.get(Number(identifier));
    }
    if (!asset) {
      asset = getAssetByCode.get(identifier);
    }
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const meter = getAssetCurrentHoursInfo(asset.id);
    return {
      ok: true,
      asset_id: asset.id,
      asset_code: asset.asset_code,
      asset_name: asset.asset_name,
      category: asset.category ?? null,
      current_hours: meter.hours,
      hour_meter_source: meter.source,
      latest_daily_work_date: meter.latest_work_date ?? null,
    };
  });

  // ------------------------------------------------------------
  // Every machine's history on one Excel sheet (the machines on the chosen site)
  // GET /api/assets/history.xlsx?start=YYYY-MM-DD&end=YYYY-MM-DD&include_archived=0|1
  // ------------------------------------------------------------
  app.get("/history.xlsx", async (req, reply) => {
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();
    if ((start && !isDate(start)) || (end && !isDate(end))) return reply.code(400).send({ error: "Dates must be YYYY-MM-DD." });
    if (start && end && start > end) return reply.code(400).send({ error: "The start date is after the end date." });
    const includeArchived = String(req.query?.include_archived || "0") === "1";
    const assets = db.prepare(`
      SELECT * FROM assets
      WHERE (? = 1 OR COALESCE(archived, 0) = 0)${andSite(null, req)}
      ORDER BY asset_code ASC
    `).all(includeArchived ? 1 : 0);
    const siteCode = String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
    const { workbook } = await assetHistoryWorkbook(assets, { start, end, siteCode });
    const period = start || end ? `_${start || "start"}_to_${end || "today"}` : "";
    reply.header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    reply.header("Content-Disposition", `attachment; filename="IRONLOG_Machine_History_${siteCode.replace(/[^a-z0-9_-]/g, "")}${period}.xlsx"`);
    return reply.send(Buffer.from(await workbook.xlsx.writeBuffer()));
  });

  // ============================================================
  // ASSET HISTORY (Timeline)
  // GET /api/assets/:asset_code/history?start=YYYY-MM-DD&end=YYYY-MM-DD
  // ============================================================
  app.get("/:asset_code/history", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    const start = String(req.query?.start || "").trim();
    const end = String(req.query?.end || "").trim();

    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const siteCode = String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
    const { summary, history } = buildAssetHistory(asset, { start, end, siteCode });

    return {
      ok: true,
      asset: {
        asset_code: asset.asset_code,
        asset_name: asset.asset_name,
        category: asset.category,
        archived: Number(asset.archived),
        archive_reason: asset.archive_reason,
        archived_at: asset.archived_at,
      },
      summary,
      history,
    };
  });
}

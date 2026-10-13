// IRONLOG/api/utils/fuelCostReport.js — Fuel cost per equipment for Reports → Fuel Cost PDF.
// Fuel cost uses the same formula as /api/reports/cost-monthly.xlsx (and buildAssetPeriodCosts):
// litres × (log unit cost, else asset fuel cost per litre, else the fuel cost default).

import { buildPdfBuffer, kvGrid, sectionTitle, table, tryDrawLogo } from "./pdfGenerator.js";

const round2 = (n) => Number(Number(n || 0).toFixed(2));

function money(v) {
  return Number(v || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function litres(v) {
  return Number(v || 0).toLocaleString("en-US", { minimumFractionDigits: 0, maximumFractionDigits: 1 });
}

function shortText(v, max) {
  const s = String(v ?? "").replace(/\s+/g, " ").trim();
  return s.length > max ? `${s.slice(0, Math.max(1, max - 1))}...` : s;
}

/**
 * Per-asset fuel cost for an inclusive YYYY-MM-DD period, plus a category roll-up and totals.
 * `defaults` is the cost_settings defaults object (fuel_cost_per_liter_default is used).
 */
export function buildFuelCostReport(db, period, defaults) {
  const fallbackRate = Number(defaults?.fuel_cost_per_liter_default ?? 1.5);
  const assetRows = db.prepare(`
    SELECT a.asset_code, a.asset_name, a.category,
      COUNT(fl.id) AS fills,
      COALESCE(SUM(fl.liters), 0) AS liters,
      COALESCE(SUM(fl.liters * COALESCE(fl.unit_cost_per_liter, a.fuel_cost_per_liter, ?)), 0) AS fuel_cost
    FROM fuel_logs fl
    JOIN assets a ON a.id = fl.asset_id
    WHERE fl.log_date BETWEEN ? AND ?
    GROUP BY a.id
  `).all(fallbackRate, period.start, period.end);

  const rows = assetRows
    .map((r) => {
      const liters = Number(r.liters || 0);
      const fuelCost = Number(r.fuel_cost || 0);
      return {
        asset_code: String(r.asset_code || ""),
        asset_name: r.asset_name || "",
        category: r.category || "Unassigned",
        fills: Number(r.fills || 0),
        liters: round2(liters),
        fuel_cost: round2(fuelCost),
        cost_per_liter: liters > 0 ? round2(fuelCost / liters) : null,
      };
    })
    .filter((r) => r.fuel_cost > 0)
    .sort((a, b) => b.fuel_cost - a.fuel_cost);

  const byCategory = new Map();
  for (const r of rows) {
    if (!byCategory.has(r.category)) byCategory.set(r.category, { category: r.category, assets: 0, liters: 0, fuel_cost: 0 });
    const c = byCategory.get(r.category);
    c.assets += 1;
    c.liters += r.liters;
    c.fuel_cost += r.fuel_cost;
  }
  const categories = Array.from(byCategory.values())
    .map((c) => ({ ...c, liters: round2(c.liters), fuel_cost: round2(c.fuel_cost), cost_per_liter: c.liters > 0 ? round2(c.fuel_cost / c.liters) : null }))
    .sort((a, b) => b.fuel_cost - a.fuel_cost);

  const totalLiters = rows.reduce((acc, r) => acc + r.liters, 0);
  const totalCost = rows.reduce((acc, r) => acc + r.fuel_cost, 0);
  const totals = {
    assets: rows.length,
    fills: rows.reduce((acc, r) => acc + r.fills, 0),
    liters: round2(totalLiters),
    fuel_cost: round2(totalCost),
    cost_per_liter: totalLiters > 0 ? round2(totalCost / totalLiters) : null,
    default_cost_per_liter: fallbackRate,
  };

  return { period, rows, categories, totals };
}

/** Render the fuel cost report as a PDF buffer (same layout helpers as the maintenance cost PDF). */
export function buildFuelCostPdf(report, { logoPath = "", db } = {}) {
  const { period, rows, categories, totals } = report;
  return buildPdfBuffer(
    (doc) => {
      if (logoPath) tryDrawLogo(doc, logoPath);
      sectionTitle(doc, "Fuel Cost Report (USD)");
      kvGrid(doc, [
        { k: "Period", v: `${period.start} to ${period.end}` },
        { k: "Equipment with fuel cost", v: String(totals.assets) },
        { k: "Fuel issues", v: String(totals.fills) },
        { k: "Litres issued", v: litres(totals.liters) },
        { k: "Fuel cost total", v: `$${money(totals.fuel_cost)}` },
        { k: "Average cost per litre", v: totals.cost_per_liter == null ? "-" : `$${money(totals.cost_per_liter)}` },
      ], 2);

      sectionTitle(doc, "By Category");
      table(
        doc,
        [
          { key: "category", label: "Category", width: 0.34 },
          { key: "assets", label: "Equipment", width: 0.14, align: "right" },
          { key: "liters", label: "Litres", width: 0.18, align: "right" },
          { key: "cost_per_liter", label: "$/L", width: 0.14, align: "right" },
          { key: "fuel_cost", label: "Fuel Cost", width: 0.20, align: "right" },
        ],
        categories.length
          ? categories.map((c) => ({
              category: shortText(c.category, 30),
              assets: String(c.assets),
              liters: litres(c.liters),
              cost_per_liter: c.cost_per_liter == null ? "-" : money(c.cost_per_liter),
              fuel_cost: money(c.fuel_cost),
            }))
          : [{ category: "No fuel cost for selected period", assets: "-", liters: "-", cost_per_liter: "-", fuel_cost: "-" }]
      );

      sectionTitle(doc, "By Equipment");
      table(
        doc,
        [
          { key: "asset_code", label: "Asset", width: 0.13 },
          { key: "asset_name", label: "Name", width: 0.24 },
          { key: "category", label: "Category", width: 0.15 },
          { key: "fills", label: "Issues", width: 0.09, align: "right" },
          { key: "liters", label: "Litres", width: 0.13, align: "right" },
          { key: "cost_per_liter", label: "$/L", width: 0.10, align: "right" },
          { key: "fuel_cost", label: "Fuel Cost", width: 0.16, align: "right" },
        ],
        rows.length
          ? rows.map((r) => ({
              asset_code: r.asset_code,
              asset_name: shortText(r.asset_name, 30),
              category: shortText(r.category, 20),
              fills: String(r.fills),
              liters: litres(r.liters),
              cost_per_liter: r.cost_per_liter == null ? "-" : money(r.cost_per_liter),
              fuel_cost: money(r.fuel_cost),
            }))
          : [{
              asset_code: "-",
              asset_name: "No fuel cost records for selected period",
              category: "",
              fills: "-",
              liters: "-",
              cost_per_liter: "-",
              fuel_cost: "-",
            }]
      );

      doc.moveDown(0.6);
      doc.fontSize(8).fillColor("#555555").text(
        `Fuel cost = litres × the cost per litre on the fuel log, else the machine's fuel rate, else the default ($${money(totals.default_cost_per_liter)}/L). Same figures as the Monthly Cost Excel.`,
        { width: doc.page.width - doc.page.margins.left - doc.page.margins.right }
      );
    },
    {
      db,
      title: "IRONLOG",
      subtitle: "Fuel Cost Report",
      rightText: `${period.start} to ${period.end}`,
      showPageNumbers: true,
    }
  );
}

// IRONLOG/api/routes/maintenance/costing.routes.js — costing gaps and Borris's
// help filling them.
//   GET  /api/maintenance/costing-gaps           open gaps (My Work card)
//   POST /api/maintenance/costing-gaps/dismiss   { key, reason } mark "not needed"
//   POST /api/maintenance/costing-gaps/assist    { key } evidence + checked proposal
// Proposals are never saved here: the person applies them through the normal
// service-cost and part-cost endpoints.
import { db } from "../../db/client.js";
import { aiChatText, aiProviderSummary, isAiConfigured } from "../../utils/ironmind.js";
import { borrisNumCtx, getLastLlmChatError } from "../../utils/llmChat.js";
import { buildDueListFromPlans } from "../../utils/serviceSchedule.js";
import { dismissCostingGap, findCostingGaps, setPlanningSnapshotProvider, summarizeCostingGaps } from "../../utils/costingGaps.js";
import {
  buildPartEvidence,
  buildServiceEvidence,
  proposalForPart,
  proposalFromAi,
  proposalFromHistory,
  serviceProposalMessages,
} from "../../utils/costingAssist.js";
import { writeAudit } from "../../utils/audit.js";

const COSTING_ROLES = ["admin", "supervisor", "workshop_admin", "plant_manager", "site_manager"];
// Services due within this many hours (or km) are checked for a price.
const HORIZON = 250;
let borrisBusy = false;

/** BORRIS_COSTING_TIMEOUT_MS, default 45 s, never above 80 s (proxies cut at ~100 s). */
export function costingTimeoutMs() {
  const n = Number(process.env.BORRIS_COSTING_TIMEOUT_MS || 45000);
  return Number.isFinite(n) && n > 0 ? Math.min(80000, n) : 45000;
}

function localToday() {
  return new Intl.DateTimeFormat("en-CA", { timeZone: "Africa/Johannesburg", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
}

export default function registerCostingRoutes(app, ctx) {
  const { buildUpcomingServiceCostForecasts, getAssetCurrentHours, requireMaintenanceRoles } = ctx;

  function upcomingForecasts() {
    const plans = db.prepare(`
      SELECT mp.*, mp.id AS plan_id, a.asset_code, a.asset_name, a.category
      FROM maintenance_plans mp JOIN assets a ON a.id = mp.asset_id
      WHERE mp.active = 1 AND a.active = 1 AND COALESCE(a.is_standby, 0) = 0 AND COALESCE(a.archived, 0) = 0
    `).all();
    const due = buildDueListFromPlans(plans, (id) => getAssetCurrentHours(id), 50);
    return buildUpcomingServiceCostForecasts(
      db,
      due.map((r) => ({ ...r, last_service_hours: r.next_due_hours - r.interval_hours })),
      { maxRemainingHours: HORIZON },
    );
  }

  // What Borris's Ask sees for planning and costing questions.
  setPlanningSnapshotProvider(() => {
    const forecasts = upcomingForecasts();
    const gaps = findCostingGaps(db, { forecasts, today: localToday() });
    return {
      upcoming_services: forecasts.slice(0, 20).map((r) => ({
        asset: r.asset_code,
        service: r.service_name,
        due_in: Number(r.remaining_hours),
        est_cost: r.forecast?.est_total_cost || 0,
        cost_source: r.forecast?.cost_source || "none",
      })),
      upcoming_total_cost: Number(forecasts.reduce((s, r) => s + Number(r.forecast?.est_total_cost || 0), 0).toFixed(2)),
      costing_gaps: { ...summarizeCostingGaps(gaps), top: gaps.slice(0, 8).map((g) => g.title) },
    };
  });

  app.get("/costing-gaps", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, COSTING_ROLES)) return;
      const gaps = findCostingGaps(db, { forecasts: upcomingForecasts(), today: localToday() });
      return { ok: true, horizon: HORIZON, summary: summarizeCostingGaps(gaps), gaps };
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/costing-gaps/dismiss", async (req, reply) => {
    if (!requireMaintenanceRoles(req, reply, COSTING_ROLES)) return;
    const key = String(req.body?.key || "").trim();
    if (!key) return reply.code(400).send({ ok: false, error: "key is required" });
    const user = String(req.headers["x-user-name"] || "");
    dismissCostingGap(db, key, { reason: req.body?.reason, user });
    writeAudit(db, req, { module: "costing", action: "costing_gap.dismiss", entity_type: "costing_gap", entity_id: key, payload: { reason: req.body?.reason || null } });
    return { ok: true, key };
  });

  app.post("/costing-gaps/assist", async (req, reply) => {
    try {
      if (!requireMaintenanceRoles(req, reply, COSTING_ROLES)) return;
      const key = String(req.body?.key || "").trim();
      const [type, ref] = [key.split(":")[0], key.slice(key.indexOf(":") + 1)];

      if (type === "part_zero_cost") {
        const evidence = buildPartEvidence(db, ref);
        if (!evidence) return reply.code(404).send({ ok: false, error: "part not found" });
        return { ok: true, key, kind: "part", proposal: proposalForPart(evidence) };
      }

      if (type !== "service_unpriced") return reply.code(400).send({ ok: false, error: "Borris can help with unpriced services and $0 parts" });
      const evidence = buildServiceEvidence(db, Number(ref));
      if (!evidence) return reply.code(404).send({ ok: false, error: "service plan not found" });

      const ai = { configured: isAiConfigured(), ...aiProviderSummary(), used: false, error: null };
      let proposal = null;
      if (ai.configured && borrisBusy) {
        ai.error = "Borris is busy with another proposal";
      } else if (ai.configured) {
        // One proposal at a time, and a hard time limit well under the ~100 s a
        // proxy allows: a local model must never tie up the server.
        borrisBusy = true;
        try {
          const text = await aiChatText(serviceProposalMessages(evidence), {
            temperature: 0,
            max_tokens: 700,
            num_ctx: borrisNumCtx(),
            json: true,
            timeout_ms: costingTimeoutMs(),
          });
          proposal = text == null ? null : proposalFromAi(db, evidence, text);
          ai.used = Boolean(proposal);
          if (!proposal) ai.error = text == null ? (getLastLlmChatError() || "no reply") : "reply was not usable JSON";
        } finally {
          borrisBusy = false;
        }
      }
      // Borris came back empty-handed (or is offline): fall back to history so
      // the person still gets a starting point, and say so.
      if (!proposal || (!proposal.lines.length && !proposal.questions.length)) {
        const fromHistory = proposalFromHistory(db, evidence);
        if (!proposal || fromHistory.lines.length) proposal = { ...fromHistory, borris_note: proposal?.summary || null };
      }
      return {
        ok: true,
        key,
        kind: "service",
        plan: evidence.plan,
        proposal,
        ai,
        evidence: {
          own_services: evidence.own_history.services,
          peer_assets: evidence.peers.assets,
          store_candidates: evidence.store_candidates.length,
          manual_passages: evidence.manual.length,
        },
      };
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });
}

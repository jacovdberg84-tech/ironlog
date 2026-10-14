// IRONLOG/api/routes/workorders/combine.routes.js — combine a machine's open pre-start work orders.
// Registered by routes/workorders.routes.js; shared helpers arrive through ctx.
import { db } from "../../db/client.js";
import { combinePrestartWorkOrders, previewCombine } from "../../utils/workOrderCombine.js";
import { writeAudit } from "../../utils/audit.js";

const LEADS = ["admin", "supervisor", "workshop_admin", "plant_manager"];

export default function registerCombineRoutes(app, ctx) {
  const { requireRoles } = ctx;

  // What combining would do: which job is kept and which are folded into it.
  app.get("/:id/combine-preview", async (req, reply) => {
    if (!requireRoles(req, reply, LEADS)) return;
    const result = previewCombine(db, Number(req.params.id));
    return reply.code(result.ok ? 200 : 400).send(result);
  });

  app.post("/:id/combine", async (req, reply) => {
    if (!requireRoles(req, reply, LEADS)) return;
    if (req.body?.confirm !== true) return reply.code(400).send({ ok: false, error: "Confirm the combine first." });
    const actor = String(req.headers["x-user-name"] || "").trim() || "workshop lead";
    const result = combinePrestartWorkOrders(db, Number(req.params.id), { actor });
    if (!result.ok) return reply.code(400).send(result);
    writeAudit(db, req, {
      module: "workorders",
      action: "combine_prestart",
      entity_type: "work_order",
      entity_id: result.keep_id,
      payload: { combined_ids: result.combined_ids, moved: result.moved },
    });
    return reply.send(result);
  });
}

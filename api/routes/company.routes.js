// IRONLOG/api/routes/company.routes.js — the customer company's profile.
//   GET /api/company/profile   any signed-in user (screens use the name, currency, time zone)
//   PUT /api/company/profile   administrators only; partial updates
import { db } from "../db/client.js";
import { ensureAuditTable, writeAudit } from "../utils/audit.js";
import { COMPANY_FIELDS, getCompanyProfile, saveCompanyProfile } from "../utils/companyProfile.js";
import { requireAnyRole } from "../utils/request.js";

export default async function companyRoutes(app) {
  ensureAuditTable(db);
  app.get("/profile", async () => ({ ok: true, profile: getCompanyProfile(db), fields: COMPANY_FIELDS }));

  app.put("/profile", async (req, reply) => {
    if (!requireAnyRole(req, reply, ["admin"])) return;
    const before = getCompanyProfile(db);
    const r = saveCompanyProfile(db, req.body || {});
    if (!r.ok) return reply.code(400).send({ ok: false, error: r.errors.join(" "), errors: r.errors });
    const changed = Object.keys(COMPANY_FIELDS).filter((f) => before[f] !== r.profile[f]);
    if (changed.length) {
      try {
        writeAudit(db, req, {
          module: "company",
          action: "profile_update",
          entity_type: "company_profile",
          entity_id: "1",
          payload: Object.fromEntries(changed.map((f) => [f, { from: before[f], to: r.profile[f] }])),
        });
      } catch {}
    }
    return r;
  });
}

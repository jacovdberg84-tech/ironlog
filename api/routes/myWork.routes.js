// IRONLOG/api/routes/myWork.routes.js — GET /api/my-work (home screen for the signed-in user)
import { db } from "../db/client.js";
import { getRoles, getSiteCode, getUser } from "../utils/request.js";
import { buildMyWork } from "../utils/myWork.js";
import { summarizeWaiting, workshopWaitingOnParts } from "../utils/partsWaiting.js";

function localToday() {
  return new Intl.DateTimeFormat("en-CA", {
    timeZone: "Africa/Johannesburg",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(new Date());
}

export default async function myWorkRoutes(app) {
  app.get("/", async (req, reply) => {
    try {
      return {
        ok: true,
        ...buildMyWork(db, {
          user: getUser(req, ""),
          roles: getRoles(req),
          site: getSiteCode(req),
          today: localToday(),
        }),
      };
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  // Everything the workshop is waiting on from stores (full list for the stores queue).
  app.get("/waiting-parts", async (req, reply) => {
    try {
      const rows = workshopWaitingOnParts(db, { site: getSiteCode(req) });
      return { ok: true, summary: summarizeWaiting(rows), rows };
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });
}

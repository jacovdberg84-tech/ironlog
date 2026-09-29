// IRONLOG/api/routes/myWork.routes.js — GET /api/my-work (home screen for the signed-in user)
import { db } from "../db/client.js";
import { getRoles, getSiteCode, getUser } from "../utils/request.js";
import { buildMyWork } from "../utils/myWork.js";

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
}

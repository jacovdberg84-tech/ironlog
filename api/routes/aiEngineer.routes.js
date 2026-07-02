import { db } from "../db/client.js";

const ALLOWED_ROLES = new Set([
  "admin",
  "supervisor",
  "plant_manager",
  "site_manager",
  "workshop_manager",
  "executive",
]);

function getSiteCode(req) {
  return String(req.headers["x-site-code"] || "main").trim().toLowerCase() || "main";
}

function getUserName(req) {
  return String(req.headers["x-user-name"] || req.headers["x-user"] || "system").trim() || "system";
}

function getRoles(req) {
  const many = String(req.headers["x-user-roles"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  const one = String(req.headers["x-user-role"] || "")
    .split(",")
    .map((x) => String(x || "").trim().toLowerCase())
    .filter(Boolean);
  return Array.from(new Set([...many, ...one]));
}

function requireRole(req, reply) {
  const roles = getRoles(req);
  if (!roles.some((r) => ALLOWED_ROLES.has(r))) {
    reply.code(403).send({ ok: false, error: "not allowed" });
    return false;
  }
  return true;
}

function ensureTables() {
  db.prepare(`
    CREATE TABLE IF NOT EXISTS ai_engineer_requests (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL DEFAULT 'main',
      title TEXT NOT NULL,
      request_text TEXT NOT NULL,
      priority TEXT NOT NULL DEFAULT 'medium',
      status TEXT NOT NULL DEFAULT 'draft',
      risk_level TEXT NOT NULL DEFAULT 'medium',
      requested_by TEXT NOT NULL DEFAULT 'system',
      approved_by TEXT,
      rejected_by TEXT,
      rejection_reason TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `).run();
  db.prepare(`
    CREATE TABLE IF NOT EXISTS ai_engineer_runs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      request_id INTEGER NOT NULL,
      run_type TEXT NOT NULL DEFAULT 'plan',
      status TEXT NOT NULL DEFAULT 'queued',
      summary TEXT,
      details_json TEXT,
      started_by TEXT NOT NULL DEFAULT 'system',
      started_at TEXT NOT NULL DEFAULT (datetime('now')),
      finished_at TEXT,
      FOREIGN KEY (request_id) REFERENCES ai_engineer_requests(id) ON DELETE CASCADE
    )
  `).run();
  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_ai_engineer_requests_site
    ON ai_engineer_requests(site_code, created_at)
  `).run();
  db.prepare(`
    CREATE INDEX IF NOT EXISTS idx_ai_engineer_runs_request
    ON ai_engineer_runs(request_id, started_at)
  `).run();
}

function isValidPriority(p) {
  return ["low", "medium", "high", "urgent"].includes(String(p || "").toLowerCase());
}

function nowSql() {
  return "datetime('now')";
}

function fetchRequestById(id, siteCode) {
  return db.prepare(`
    SELECT *
    FROM ai_engineer_requests
    WHERE id = ?
      AND LOWER(TRIM(COALESCE(site_code, 'main'))) = ?
    LIMIT 1
  `).get(id, siteCode);
}

function mapRequest(r) {
  if (!r) return null;
  return {
    id: Number(r.id || 0),
    site_code: String(r.site_code || "main"),
    title: String(r.title || ""),
    request_text: String(r.request_text || ""),
    priority: String(r.priority || "medium"),
    status: String(r.status || "draft"),
    risk_level: String(r.risk_level || "medium"),
    requested_by: String(r.requested_by || ""),
    approved_by: r.approved_by ? String(r.approved_by) : null,
    rejected_by: r.rejected_by ? String(r.rejected_by) : null,
    rejection_reason: r.rejection_reason ? String(r.rejection_reason) : null,
    created_at: String(r.created_at || ""),
    updated_at: String(r.updated_at || ""),
  };
}

function mapRun(r) {
  if (!r) return null;
  let details = null;
  try {
    details = r.details_json ? JSON.parse(String(r.details_json)) : null;
  } catch {
    details = null;
  }
  return {
    id: Number(r.id || 0),
    request_id: Number(r.request_id || 0),
    run_type: String(r.run_type || "plan"),
    status: String(r.status || "queued"),
    summary: r.summary ? String(r.summary) : "",
    details,
    started_by: String(r.started_by || ""),
    started_at: String(r.started_at || ""),
    finished_at: r.finished_at ? String(r.finished_at) : null,
  };
}

export default async function aiEngineerRoutes(app) {
  ensureTables();

  app.get("/requests", async (req, reply) => {
    if (!requireRole(req, reply)) return;
    const siteCode = getSiteCode(req);
    const status = String(req.query?.status || "").trim().toLowerCase();
    const where = [
      `LOWER(TRIM(COALESCE(site_code, 'main'))) = ?`,
    ];
    const params = [siteCode];
    if (status) {
      where.push("LOWER(TRIM(COALESCE(status, 'draft'))) = ?");
      params.push(status);
    }
    const rows = db.prepare(`
      SELECT *
      FROM ai_engineer_requests
      WHERE ${where.join(" AND ")}
      ORDER BY created_at DESC, id DESC
      LIMIT 200
    `).all(...params);
    return reply.send({ ok: true, rows: rows.map(mapRequest) });
  });

  app.get("/requests/:id", async (req, reply) => {
    if (!requireRole(req, reply)) return;
    const siteCode = getSiteCode(req);
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) {
      return reply.code(400).send({ ok: false, error: "valid id required" });
    }
    const requestRow = fetchRequestById(id, siteCode);
    if (!requestRow) {
      return reply.code(404).send({ ok: false, error: "request not found" });
    }
    const runs = db.prepare(`
      SELECT *
      FROM ai_engineer_runs
      WHERE request_id = ?
      ORDER BY started_at DESC, id DESC
      LIMIT 100
    `).all(id).map(mapRun);
    return reply.send({ ok: true, row: mapRequest(requestRow), runs });
  });

  app.post("/requests", async (req, reply) => {
    if (!requireRole(req, reply)) return;
    const siteCode = getSiteCode(req);
    const user = getUserName(req);
    const title = String(req.body?.title || "").trim();
    const requestText = String(req.body?.request_text || "").trim();
    const priorityRaw = String(req.body?.priority || "medium").trim().toLowerCase();
    const priority = isValidPriority(priorityRaw) ? priorityRaw : "medium";
    if (!title) return reply.code(400).send({ ok: false, error: "title required" });
    if (!requestText) return reply.code(400).send({ ok: false, error: "request_text required" });

    const risk = priority === "urgent" || priority === "high" ? "high" : "medium";
    const result = db.prepare(`
      INSERT INTO ai_engineer_requests (
        site_code, title, request_text, priority, status, risk_level, requested_by, created_at, updated_at
      ) VALUES (?, ?, ?, ?, 'draft', ?, ?, ${nowSql()}, ${nowSql()})
    `).run(siteCode, title, requestText, priority, risk, user);

    const row = fetchRequestById(Number(result.lastInsertRowid || 0), siteCode);
    return reply.send({ ok: true, row: mapRequest(row) });
  });

  app.post("/requests/:id/plan", async (req, reply) => {
    if (!requireRole(req, reply)) return;
    const siteCode = getSiteCode(req);
    const user = getUserName(req);
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ ok: false, error: "valid id required" });
    const row = fetchRequestById(id, siteCode);
    if (!row) return reply.code(404).send({ ok: false, error: "request not found" });

    const details = {
      phase: "mvp-plan",
      scope: "planning",
      next_steps: [
        "Analyze impacted files",
        "Generate implementation diff",
        "Run lint/test checks",
        "Require approval before deploy",
      ],
      note: "This is a foundation planner run. Live autonomous code generation is wired next.",
    };

    const runInfo = db.prepare(`
      INSERT INTO ai_engineer_runs (
        request_id, run_type, status, summary, details_json, started_by, started_at, finished_at
      ) VALUES (?, 'plan', 'completed', ?, ?, ?, ${nowSql()}, ${nowSql()})
    `).run(
      id,
      "MVP plan created. Ready for execution wiring.",
      JSON.stringify(details),
      user,
    );

    db.prepare(`
      UPDATE ai_engineer_requests
      SET status = 'planned',
          updated_at = ${nowSql()}
      WHERE id = ?
    `).run(id);

    const updated = fetchRequestById(id, siteCode);
    const run = db.prepare(`SELECT * FROM ai_engineer_runs WHERE id = ?`).get(runInfo.lastInsertRowid);
    return reply.send({ ok: true, row: mapRequest(updated), run: mapRun(run) });
  });

  app.post("/requests/:id/approve", async (req, reply) => {
    if (!requireRole(req, reply)) return;
    const siteCode = getSiteCode(req);
    const user = getUserName(req);
    const id = Number(req.params?.id || 0);
    const row = fetchRequestById(id, siteCode);
    if (!row) return reply.code(404).send({ ok: false, error: "request not found" });

    db.prepare(`
      UPDATE ai_engineer_requests
      SET status = 'approved',
          approved_by = ?,
          rejected_by = NULL,
          rejection_reason = NULL,
          updated_at = ${nowSql()}
      WHERE id = ?
    `).run(user, id);

    const updated = fetchRequestById(id, siteCode);
    return reply.send({ ok: true, row: mapRequest(updated) });
  });

  app.post("/requests/:id/reject", async (req, reply) => {
    if (!requireRole(req, reply)) return;
    const siteCode = getSiteCode(req);
    const user = getUserName(req);
    const id = Number(req.params?.id || 0);
    const reason = String(req.body?.reason || "").trim();
    const row = fetchRequestById(id, siteCode);
    if (!row) return reply.code(404).send({ ok: false, error: "request not found" });
    if (!reason) return reply.code(400).send({ ok: false, error: "reason required" });

    db.prepare(`
      UPDATE ai_engineer_requests
      SET status = 'rejected',
          rejected_by = ?,
          rejection_reason = ?,
          updated_at = ${nowSql()}
      WHERE id = ?
    `).run(user, reason, id);
    const updated = fetchRequestById(id, siteCode);
    return reply.send({ ok: true, row: mapRequest(updated) });
  });
}

import { db } from "../db/client.js";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";
import OpenAI from "openai";

const ALLOWED_ROLES = new Set([
  "admin",
  "supervisor",
  "plant_manager",
  "site_manager",
  "workshop_manager",
  "executive",
]);

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(__dirname, "..", "..");
const WORKTREE_BASE = path.join(REPO_ROOT, ".ai-engineer", "worktrees");
const AI_ENGINEER_MODEL = process.env.AI_ENGINEER_MODEL || "gpt-4o-mini";
let openaiClient = null;

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

function slug(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 40) || "request";
}

function runCmd(command, args, cwd, timeoutMs = 120000) {
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    timeout: timeoutMs,
    shell: false,
    windowsHide: true,
  });
  return {
    command: `${command} ${args.join(" ")}`,
    cwd,
    exit_code: typeof result.status === "number" ? result.status : 1,
    stdout: String(result.stdout || "").trim(),
    stderr: String(result.stderr || "").trim(),
    timed_out: Boolean(result.error && String(result.error.message || "").toLowerCase().includes("timed out")),
  };
}

function ensureDir(p) {
  try {
    fs.mkdirSync(p, { recursive: true });
  } catch {}
}

function getOpenAIClient() {
  const apiKey = String(process.env.OPENAI_API_KEY || "").trim();
  if (!apiKey) return null;
  if (!openaiClient) openaiClient = new OpenAI({ apiKey });
  return openaiClient;
}

function extractUnifiedDiff(text) {
  const s = String(text || "");
  const fenced = s.match(/```diff\s*([\s\S]*?)```/i);
  if (fenced?.[1]) return fenced[1].trim();
  const fallback = s.match(/(diff --git[\s\S]*)/i);
  return fallback?.[1]?.trim() || "";
}

function readFileSafe(p, maxChars = 16000) {
  try {
    const content = fs.readFileSync(p, "utf8");
    return content.length > maxChars ? `${content.slice(0, maxChars)}\n/* truncated */` : content;
  } catch {
    return "";
  }
}

async function generateUnifiedDiffViaModel({ requestRow, worktreePath, targetFiles }) {
  const client = getOpenAIClient();
  if (!client) {
    return {
      ok: false,
      error: "OPENAI_API_KEY is not configured on the API service.",
      model: AI_ENGINEER_MODEL,
    };
  }

  const existing = targetFiles
    .map((rel) => {
      const abs = path.join(worktreePath, rel);
      const body = readFileSafe(abs);
      return body
        ? `FILE: ${rel}\n-----\n${body}\n-----\n`
        : `FILE: ${rel}\n-----\n(unreadable or missing)\n-----\n`;
    })
    .join("\n");

  const systemPrompt = [
    "You are a senior software engineer editing an existing JavaScript/HTML codebase.",
    "Output ONLY a unified diff patch in git format starting with 'diff --git'.",
    "Rules:",
    "- Wrap output in a single ```diff fenced block.",
    "- Limit edits to files provided in TARGET FILES.",
    "- Keep changes minimal and backward compatible.",
    "- Ensure code remains syntactically valid.",
  ].join("\n");

  const userPrompt = [
    `REQUEST ID: ${Number(requestRow?.id || 0)}`,
    `TITLE: ${String(requestRow?.title || "")}`,
    `REQUEST: ${String(requestRow?.request_text || "")}`,
    "",
    "TARGET FILES:",
    ...targetFiles.map((f) => `- ${f}`),
    "",
    "CURRENT FILE CONTENTS:",
    existing,
    "",
    "Produce a unified diff now.",
  ].join("\n");

  const resp = await client.responses.create({
    model: AI_ENGINEER_MODEL,
    input: [
      { role: "system", content: [{ type: "input_text", text: systemPrompt }] },
      { role: "user", content: [{ type: "input_text", text: userPrompt }] },
    ],
  });

  const raw = String(resp.output_text || "").trim();
  const diff = extractUnifiedDiff(raw);
  if (!diff) {
    return {
      ok: false,
      error: "Model returned no unified diff.",
      model: AI_ENGINEER_MODEL,
      raw_preview: raw.slice(0, 2000),
    };
  }
  return { ok: true, model: AI_ENGINEER_MODEL, diff, raw_preview: raw.slice(0, 2000) };
}

function chooseTargetFiles(title, text) {
  const hay = `${String(title || "")} ${String(text || "")}`.toLowerCase();
  const files = new Set();
  if (/maint|service|mechanic|workshop/.test(hay)) {
    files.add("web/maintenance.html");
    files.add("web/maintenance.js");
    files.add("api/routes/maintenance.routes.js");
  }
  if (/lube|oil|dashboard/.test(hay)) {
    files.add("web/app.js");
    files.add("api/routes/dashboard.routes.js");
    files.add("api/routes/reports.routes.js");
  }
  if (/work.?order|breakdown/.test(hay)) {
    files.add("api/routes/workorders.routes.js");
    files.add("api/routes/breakdowns.routes.js");
  }
  if (/api|endpoint|route/.test(hay)) {
    files.add("api/server.js");
  }
  if (!files.size) {
    files.add("web/app.js");
    files.add("api/routes/maintenance.routes.js");
  }
  return [...files];
}

function buildProposalMarkdown(requestRow, targetFiles) {
  const title = String(requestRow?.title || "").trim();
  const requestText = String(requestRow?.request_text || "").trim();
  const pri = String(requestRow?.priority || "medium");
  const risk = String(requestRow?.risk_level || "medium");
  const bullets = [
    "Implement minimal, backwards-compatible changes.",
    "Prefer additive API changes and non-breaking UI defaults.",
    "Add or adjust tests/verification where feasible.",
    "Keep modified files constrained to target scope.",
  ];
  return [
    `# AI Engineer Proposal`,
    ``,
    `## Request`,
    `- ID: ${Number(requestRow?.id || 0)}`,
    `- Title: ${title}`,
    `- Priority: ${pri}`,
    `- Risk: ${risk}`,
    ``,
    `### User request`,
    requestText || "(empty)",
    ``,
    `## Target files`,
    ...targetFiles.map((f) => `- ${f}`),
    ``,
    `## Planned implementation`,
    ...bullets.map((b) => `- ${b}`),
    ``,
    `## Validation gates`,
    `- API import validation`,
    `- Git diff produced in isolated worktree`,
    `- Optional test/lint command execution if available`,
    ``,
  ].join("\n");
}

function getLatestExecuteRun(requestId) {
  return db.prepare(`
    SELECT *
    FROM ai_engineer_runs
    WHERE request_id = ?
      AND run_type = 'execute'
      AND status = 'completed'
    ORDER BY started_at DESC, id DESC
    LIMIT 1
  `).get(requestId);
}

function parseDetailsJson(raw) {
  try {
    return raw ? JSON.parse(String(raw)) : null;
  } catch {
    return null;
  }
}

function buildGeneratedPatchPayload({ requestRow, targetFiles, model, diffFile }) {
  return {
    schema: "ironlog-ai-engineer/patch-v2",
    request_id: Number(requestRow?.id || 0),
    title: String(requestRow?.title || ""),
    request_text: String(requestRow?.request_text || ""),
    target_files: targetFiles,
    generated_at: new Date().toISOString(),
    model: model || null,
    diff_file: diffFile || null,
    changes: targetFiles.map((f) => ({
      file: f,
      action: "modify",
      description: `AI-generated code change for request #${Number(requestRow?.id || 0)}`,
    })),
  };
}

async function generatePatchInWorktree({ requestRow, executeDetails }) {
  const worktreePath = String(executeDetails?.worktree_path || "").trim();
  if (!worktreePath) {
    return { ok: false, summary: "Missing worktree path in execute run details.", details: { execute_details: executeDetails } };
  }
  ensureDir(path.join(worktreePath, ".ai-engineer", "patches"));
  const targetFiles = Array.isArray(executeDetails?.proposal?.target_files)
    ? executeDetails.proposal.target_files
    : chooseTargetFiles(requestRow?.title, requestRow?.request_text);

  const modelResult = await generateUnifiedDiffViaModel({
    requestRow,
    worktreePath,
    targetFiles,
  });
  if (!modelResult.ok) {
    return {
      ok: false,
      summary: modelResult.error || "AI patch generation failed.",
      details: {
        worktree_path: worktreePath,
        target_files: targetFiles,
        model: modelResult.model,
        raw_preview: modelResult.raw_preview || null,
      },
    };
  }

  const reqId = Number(requestRow?.id || 0);
  const diffRel = `.ai-engineer/patches/request-${reqId}.patch`;
  const diffPath = path.join(worktreePath, diffRel);
  const metaRel = `.ai-engineer/patches/request-${reqId}.json`;
  const metaPath = path.join(worktreePath, metaRel);

  try {
    fs.writeFileSync(diffPath, `${modelResult.diff}\n`, "utf8");
    const payload = buildGeneratedPatchPayload({
      requestRow,
      targetFiles,
      model: modelResult.model,
      diffFile: diffRel,
    });
    fs.writeFileSync(metaPath, JSON.stringify(payload, null, 2), "utf8");
  } catch (err) {
    return {
      ok: false,
      summary: "Failed to write generated patch files.",
      details: { error: String(err?.message || err) },
    };
  }

  const checkStep = runCmd("git", ["apply", "--check", diffRel], worktreePath);
  if (checkStep.exit_code !== 0) {
    return {
      ok: false,
      summary: "Generated diff failed git apply --check. Review raw output and retry.",
      details: {
        worktree_path: worktreePath,
        patch_file: diffRel,
        patch_meta_file: metaRel,
        model: modelResult.model,
        diff_preview: String(modelResult.diff || "").slice(0, 12000),
        raw_preview: modelResult.raw_preview || null,
        steps: [checkStep],
      },
    };
  }

  return {
    ok: true,
    summary: `Generated real code diff via ${modelResult.model}.`,
    details: {
      worktree_path: worktreePath,
      patch_file: diffRel,
      patch_meta_file: metaRel,
      payload: buildGeneratedPatchPayload({
        requestRow,
        targetFiles,
        model: modelResult.model,
        diffFile: diffRel,
      }),
      model: modelResult.model,
      diff_preview: String(modelResult.diff || "").slice(0, 12000),
      raw_preview: modelResult.raw_preview || null,
      steps: [checkStep],
    },
  };
}

function applyGeneratedPatchInWorktree({ requestRow, executeDetails, patchDetails }) {
  const worktreePath = String(executeDetails?.worktree_path || "").trim();
  const diffRel = String(patchDetails?.patch_file || patchDetails?.payload?.diff_file || "").trim();
  if (!worktreePath || !diffRel) {
    return {
      ok: false,
      summary: "Patch apply failed: missing worktree or patch diff file.",
      details: { worktree_path: worktreePath, patch_file: diffRel, patch_details: patchDetails },
    };
  }

  const diffPath = path.join(worktreePath, diffRel);
  if (!fs.existsSync(diffPath)) {
    return {
      ok: false,
      summary: "Patch diff file does not exist in worktree.",
      details: { patch_file: diffRel },
    };
  }

  const steps = [];
  const checkStep = runCmd("git", ["apply", "--check", diffRel], worktreePath);
  steps.push(checkStep);
  if (checkStep.exit_code !== 0) {
    return {
      ok: false,
      summary: "Patch failed git apply --check. Review diff for conflicts.",
      details: {
        worktree_path: worktreePath,
        patch_file: diffRel,
        steps,
        diff_preview: readFileSafe(diffPath, 12000),
      },
    };
  }

  const applyStep = runCmd("git", ["apply", diffRel], worktreePath);
  steps.push(applyStep);
  if (applyStep.exit_code !== 0) {
    return {
      ok: false,
      summary: "Patch apply failed during git apply.",
      details: { worktree_path: worktreePath, patch_file: diffRel, steps },
    };
  }

  ensureDir(path.join(worktreePath, ".ai-engineer", "applied"));
  const summaryPath = path.join(worktreePath, ".ai-engineer", "applied", "request-summary.md");
  const summaryEntry = [
    `## Request #${Number(requestRow?.id || 0)} - ${String(requestRow?.title || "")}`,
    `- Applied at: ${new Date().toISOString()}`,
    `- Priority: ${String(requestRow?.priority || "medium")}`,
    `- Model: ${String(patchDetails?.model || patchDetails?.payload?.model || "unknown")}`,
    `- Patch file: ${diffRel}`,
    ``,
  ].join("\n");
  try {
    fs.appendFileSync(summaryPath, `${summaryEntry}\n`, "utf8");
  } catch {}

  const statusStep = runCmd("git", ["status", "--porcelain"], worktreePath);
  const diffStep = runCmd("git", ["diff"], worktreePath);
  const apiImport = runCmd(
    "node",
    ["-e", "import('./server.js').then(()=>console.log('server-import-ok'))"],
    path.join(worktreePath, "api"),
  );
  steps.push(statusStep, diffStep, apiImport);

  const changed = String(statusStep.stdout || "")
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean);

  const ok = applyStep.exit_code === 0 && apiImport.exit_code === 0 && changed.length > 0;
  return {
    ok,
    summary: ok
      ? "AI-generated patch applied in isolated worktree. Review diff before merge/deploy."
      : "Patch apply completed with issues; inspect run details.",
    details: {
      worktree_path: worktreePath,
      patch_file: diffRel,
      changed_files: changed,
      diff_preview: String(diffStep.stdout || "").slice(0, 12000),
      gates: {
        git_apply: applyStep.exit_code === 0,
        server_import: apiImport.exit_code === 0,
        has_changes: changed.length > 0,
      },
      steps,
    },
  };
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

function latestRunForRequest(requestId) {
  return db.prepare(`
    SELECT *
    FROM ai_engineer_runs
    WHERE request_id = ?
    ORDER BY started_at DESC, id DESC
    LIMIT 1
  `).get(requestId);
}

function executeGuardedPipeline({ requestRow, startedBy }) {
  ensureDir(WORKTREE_BASE);
  const reqId = Number(requestRow?.id || 0);
  const stamp = new Date().toISOString().replace(/[-:.TZ]/g, "").slice(0, 14);
  const branch = `aie/${reqId}-${stamp}-${slug(requestRow?.title)}`;
  const worktreePath = path.join(WORKTREE_BASE, `req-${reqId}-${stamp}`);
  ensureDir(path.dirname(worktreePath));

  const steps = [];
  steps.push(runCmd("git", ["worktree", "add", "-b", branch, worktreePath, "HEAD"], REPO_ROOT));
  if (steps.at(-1)?.exit_code !== 0) {
    return {
      ok: false,
      summary: "Execution stopped: unable to create isolated worktree.",
      details: {
        branch,
        worktree_path: worktreePath,
        steps,
      },
    };
  }

  const apiDir = path.join(worktreePath, "api");
  const checks = [];
  const proposalDir = path.join(worktreePath, ".ai-engineer", "proposals");
  ensureDir(proposalDir);
  const targetFiles = chooseTargetFiles(requestRow?.title, requestRow?.request_text);
  const proposalMarkdown = buildProposalMarkdown(requestRow, targetFiles);
  const proposalPath = path.join(proposalDir, `request-${reqId}.md`);
  const proposalJsonPath = path.join(proposalDir, `request-${reqId}.json`);
  try {
    fs.writeFileSync(proposalPath, proposalMarkdown, "utf8");
    fs.writeFileSync(
      proposalJsonPath,
      JSON.stringify(
        {
          request_id: reqId,
          title: String(requestRow?.title || ""),
          request_text: String(requestRow?.request_text || ""),
          target_files: targetFiles,
          generated_at: new Date().toISOString(),
        },
        null,
        2,
      ),
      "utf8",
    );
  } catch (err) {
    checks.push({
      command: "write-proposal-files",
      cwd: proposalDir,
      exit_code: 1,
      stdout: "",
      stderr: String(err?.message || err || "proposal write failed"),
      timed_out: false,
    });
  }

  checks.push(runCmd("node", ["-e", "import('./server.js').then(()=>console.log('server-import-ok'))"], apiDir));
  checks.push(runCmd("git", ["status", "--porcelain"], worktreePath));
  checks.push(runCmd("git", ["diff", "--", ".ai-engineer/proposals"], worktreePath));

  const allOk = checks.every((s) => s.exit_code === 0);
  const serverImportStep = checks.find((s) => String(s.command || "").startsWith("node "));
  const statusStep = checks.find((s) => String(s.command || "").startsWith("git status"));
  const diffStep = checks.find((s) => String(s.command || "").startsWith("git diff"));
  const changed = String(statusStep?.stdout || "")
    .split(/\r?\n/)
    .map((x) => x.trim())
    .filter(Boolean);

  const executionPlan = {
    request_id: reqId,
    branch,
    worktree_path: worktreePath,
    started_by: startedBy,
    gates: {
      server_import: Number(serverImportStep?.exit_code || 1) === 0,
      git_diff_available: Number(diffStep?.exit_code || 1) === 0,
      has_changes: changed.length > 0,
    },
    changed_files: changed,
    proposal: {
      file: path.relative(worktreePath, proposalPath).replace(/\\/g, "/"),
      json_file: path.relative(worktreePath, proposalJsonPath).replace(/\\/g, "/"),
      target_files: targetFiles,
      markdown_preview: proposalMarkdown.slice(0, 1800),
      diff_preview: String(diffStep?.stdout || "").slice(0, 6000),
    },
    notes: [
      "Phase 3 start: execution now generates concrete implementation proposal artifacts.",
      "Next step: replace proposal generator with live code generator and apply patch stage.",
    ],
    steps: [...steps, ...checks],
  };

  return {
    ok: allOk,
    summary: allOk
      ? "Execution pipeline completed with generated proposal + diff preview in isolated worktree."
      : "Execution pipeline completed with failing gates. Review run details.",
    details: executionPlan,
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

  app.post("/requests/:id/execute", async (req, reply) => {
    if (!requireRole(req, reply)) return;
    const siteCode = getSiteCode(req);
    const user = getUserName(req);
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ ok: false, error: "valid id required" });
    const row = fetchRequestById(id, siteCode);
    if (!row) return reply.code(404).send({ ok: false, error: "request not found" });
    const st = String(row.status || "draft").toLowerCase();
    if (!["approved", "planned", "executed"].includes(st)) {
      return reply.code(400).send({ ok: false, error: "request must be planned or approved before execute" });
    }

    const runStart = db.prepare(`
      INSERT INTO ai_engineer_runs (
        request_id, run_type, status, summary, details_json, started_by, started_at
      ) VALUES (?, 'execute', 'running', ?, ?, ?, ${nowSql()})
    `).run(id, "Execution started.", JSON.stringify({ stage: "starting" }), user);
    const runId = Number(runStart.lastInsertRowid || 0);

    const result = executeGuardedPipeline({ requestRow: row, startedBy: user });
    const runStatus = result.ok ? "completed" : "failed";
    db.prepare(`
      UPDATE ai_engineer_runs
      SET status = ?,
          summary = ?,
          details_json = ?,
          finished_at = ${nowSql()}
      WHERE id = ?
    `).run(runStatus, result.summary, JSON.stringify(result.details), runId);

    db.prepare(`
      UPDATE ai_engineer_requests
      SET status = ?,
          updated_at = ${nowSql()}
      WHERE id = ?
    `).run(result.ok ? "executed" : "planned", id);

    const updated = fetchRequestById(id, siteCode);
    const run = db.prepare(`SELECT * FROM ai_engineer_runs WHERE id = ?`).get(runId);
    return reply.send({ ok: result.ok, row: mapRequest(updated), run: mapRun(run) });
  });

  app.post("/requests/:id/generate-patch", async (req, reply) => {
    if (!requireRole(req, reply)) return;
    const siteCode = getSiteCode(req);
    const user = getUserName(req);
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ ok: false, error: "valid id required" });
    const row = fetchRequestById(id, siteCode);
    if (!row) return reply.code(404).send({ ok: false, error: "request not found" });
    if (!row.approved_by) return reply.code(400).send({ ok: false, error: "request must be approved first" });

    const executeRun = getLatestExecuteRun(id);
    if (!executeRun) return reply.code(400).send({ ok: false, error: "execute run required before generate-patch" });
    const executeDetails = parseDetailsJson(executeRun.details_json) || {};

    const runStart = db.prepare(`
      INSERT INTO ai_engineer_runs (
        request_id, run_type, status, summary, details_json, started_by, started_at
      ) VALUES (?, 'generate_patch', 'running', ?, ?, ?, ${nowSql()})
    `).run(id, "Generating code patch via AI.", JSON.stringify({ stage: "starting" }), user);
    const runId = Number(runStart.lastInsertRowid || 0);

    let result;
    try {
      result = await generatePatchInWorktree({ requestRow: row, executeDetails });
    } catch (err) {
      result = {
        ok: false,
        summary: `AI patch generation error: ${String(err?.message || err)}`,
        details: { error: String(err?.message || err) },
      };
    }
    db.prepare(`
      UPDATE ai_engineer_runs
      SET status = ?,
          summary = ?,
          details_json = ?,
          finished_at = ${nowSql()}
      WHERE id = ?
    `).run(result.ok ? "completed" : "failed", result.summary, JSON.stringify(result.details), runId);
    db.prepare(`
      UPDATE ai_engineer_requests
      SET status = ?,
          updated_at = ${nowSql()}
      WHERE id = ?
    `).run(result.ok ? "patch_ready" : String(row.status || "executed"), id);

    const updated = fetchRequestById(id, siteCode);
    const run = db.prepare(`SELECT * FROM ai_engineer_runs WHERE id = ?`).get(runId);
    return reply.send({ ok: result.ok, row: mapRequest(updated), run: mapRun(run) });
  });

  app.post("/requests/:id/apply-patch", async (req, reply) => {
    if (!requireRole(req, reply)) return;
    const siteCode = getSiteCode(req);
    const user = getUserName(req);
    const id = Number(req.params?.id || 0);
    if (!Number.isFinite(id) || id <= 0) return reply.code(400).send({ ok: false, error: "valid id required" });
    const row = fetchRequestById(id, siteCode);
    if (!row) return reply.code(404).send({ ok: false, error: "request not found" });
    if (!row.approved_by) return reply.code(400).send({ ok: false, error: "request must be approved first" });

    const executeRun = getLatestExecuteRun(id);
    if (!executeRun) return reply.code(400).send({ ok: false, error: "execute run required before apply-patch" });
    const executeDetails = parseDetailsJson(executeRun.details_json) || {};
    const patchRun = db.prepare(`
      SELECT *
      FROM ai_engineer_runs
      WHERE request_id = ?
        AND run_type = 'generate_patch'
        AND status = 'completed'
      ORDER BY started_at DESC, id DESC
      LIMIT 1
    `).get(id);
    if (!patchRun) return reply.code(400).send({ ok: false, error: "generate-patch run required before apply-patch" });
    const patchDetails = parseDetailsJson(patchRun.details_json) || {};

    const runStart = db.prepare(`
      INSERT INTO ai_engineer_runs (
        request_id, run_type, status, summary, details_json, started_by, started_at
      ) VALUES (?, 'apply_patch', 'running', ?, ?, ?, ${nowSql()})
    `).run(id, "Applying generated patch.", JSON.stringify({ stage: "starting" }), user);
    const runId = Number(runStart.lastInsertRowid || 0);

    const result = applyGeneratedPatchInWorktree({
      requestRow: row,
      executeDetails,
      patchDetails,
    });
    db.prepare(`
      UPDATE ai_engineer_runs
      SET status = ?,
          summary = ?,
          details_json = ?,
          finished_at = ${nowSql()}
      WHERE id = ?
    `).run(result.ok ? "completed" : "failed", result.summary, JSON.stringify(result.details), runId);
    db.prepare(`
      UPDATE ai_engineer_requests
      SET status = ?,
          updated_at = ${nowSql()}
      WHERE id = ?
    `).run(result.ok ? "patch_applied" : String(row.status || "patch_ready"), id);

    const updated = fetchRequestById(id, siteCode);
    const run = db.prepare(`SELECT * FROM ai_engineer_runs WHERE id = ?`).get(runId);
    return reply.send({ ok: result.ok, row: mapRequest(updated), run: mapRun(run) });
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

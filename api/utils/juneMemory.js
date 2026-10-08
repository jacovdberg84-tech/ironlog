// IRONLOG/api/utils/juneMemory.js — June remembers recent conversations.
//
// Each live session starts fresh at OpenAI, so the browser saves the turns
// (what the administrator said, June's replies, tools she used) and the next
// session gets a short recap in its instructions. Kept per site and user;
// the administrator can clear it.
import { db as defaultDb } from "../db/client.js";

const KEEP_DAYS = 14;
const MAX_TURN_CHARS = 600;

function ensureTables(dbConn) {
  dbConn.exec(`
    CREATE TABLE IF NOT EXISTS june_memory_turns (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      site_code TEXT NOT NULL,
      username TEXT NOT NULL,
      role TEXT NOT NULL,
      text TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_june_memory_user ON june_memory_turns(site_code, username, id);
  `);
}

const key = ({ siteCode, user }) => ({ site: String(siteCode || "main").toLowerCase(), user: String(user || "").toLowerCase() });
const clean = (t) => String(t || "").replace(/\s+/g, " ").trim().slice(0, MAX_TURN_CHARS);

export function saveJuneTurns(context, turns, { dbConn = defaultDb } = {}) {
  ensureTables(dbConn);
  const { site, user } = key(context);
  if (!user) return 0;
  const add = dbConn.prepare(`INSERT INTO june_memory_turns (site_code, username, role, text) VALUES (?, ?, ?, ?)`);
  let n = 0;
  for (const t of (Array.isArray(turns) ? turns : []).slice(0, 20)) {
    const role = ["user", "june", "tool"].includes(t?.role) ? t.role : null;
    const text = clean(t?.text);
    if (!role || !text) continue;
    add.run(site, user, role, text);
    n += 1;
  }
  // Old turns are dropped; memory is for continuity, not an archive.
  dbConn.prepare(`DELETE FROM june_memory_turns WHERE created_at < datetime('now', ?)`).run(`-${KEEP_DAYS} days`);
  return n;
}

export function recentJuneTurns(context, { maxTurns = 40, maxChars = 6000, dbConn = defaultDb } = {}) {
  ensureTables(dbConn);
  const { site, user } = key(context);
  const rows = dbConn.prepare(`
    SELECT role, text, created_at FROM june_memory_turns
    WHERE site_code = ? AND username = ? AND created_at >= datetime('now', ?)
    ORDER BY id DESC LIMIT ?
  `).all(site, user, `-${KEEP_DAYS} days`, maxTurns);
  const out = [];
  let chars = 0;
  for (const r of rows) {
    chars += r.text.length + 30;
    if (chars > maxChars) break;
    out.push(r);
  }
  return out.reverse();
}

export function clearJuneMemory(context, { dbConn = defaultDb } = {}) {
  ensureTables(dbConn);
  const { site, user } = key(context);
  return dbConn.prepare(`DELETE FROM june_memory_turns WHERE site_code = ? AND username = ?`).run(site, user).changes;
}

/** The recap added to June's instructions at the start of a live session. */
export function juneMemoryInstructions(context, { name = "the administrator", dbConn = defaultDb } = {}) {
  const turns = recentJuneTurns(context, { dbConn });
  if (!turns.length) return "";
  const who = { user: name, june: "June", tool: "June used" };
  const lines = turns.map((t) => `[${String(t.created_at).slice(5, 16)}] ${who[t.role]}: ${t.text}`);
  return [
    `Memory of your recent conversations with ${name} (oldest first, times UTC). Use it for continuity: pick up open threads, remember what was asked and decided, and do not make ${name} repeat things.`,
    "Do not recite this memory unless asked. Facts in it may be outdated; check live Ironlog data before relying on numbers.",
    ...lines,
  ].join("\n");
}

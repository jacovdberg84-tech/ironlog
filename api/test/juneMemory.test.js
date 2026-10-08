import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-june-memory-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const { db } = await import("../db/client.js");
const m = await import("../utils/juneMemory.js");
process.on("exit", () => { try { db.close(); } catch {} rmSync(tempDir, { recursive: true, force: true }); });

test("June keeps recent turns per user and recaps them for the next session", () => {
  const jaco = { siteCode: "main", user: "Jaco" };
  const other = { siteCode: "main", user: "maria" };
  assert.equal(m.juneMemoryInstructions(jaco, { name: "Jaco" }), "", "no memory yet");
  m.saveJuneTurns(jaco, [
    { role: "user", text: "Draft the A303AM service schedule for next month" },
    { role: "tool", text: "june_maintenance_schedule_draft A303AM" },
    { role: "june", text: "Done, it's ready to download. Want the hoses added too?" },
    { role: "hacker", text: "ignored" },
    { role: "user", text: "   " },
  ]);
  m.saveJuneTurns(other, [{ role: "user", text: "Maria's private chat" }]);
  const recap = m.juneMemoryInstructions(jaco, { name: "Jaco" });
  assert.match(recap, /Jaco: Draft the A303AM service schedule/);
  assert.match(recap, /June used: june_maintenance_schedule_draft A303AM/);
  assert.match(recap, /June: Done, it's ready/);
  assert.doesNotMatch(recap, /Maria's private chat|ignored/);
  assert.equal(m.recentJuneTurns(jaco).length, 3);
  // Long histories are trimmed to the newest turns.
  m.saveJuneTurns(jaco, Array.from({ length: 20 }, (_, i) => ({ role: "user", text: `note ${i} ${"x".repeat(500)}` })));
  const trimmed = m.recentJuneTurns(jaco, { maxChars: 2000 });
  assert.ok(trimmed.length < 10 && trimmed.at(-1).text.startsWith("note 19"), "keeps the newest");
  assert.ok(m.clearJuneMemory(jaco) > 0);
  assert.equal(m.recentJuneTurns(jaco).length, 0);
  assert.equal(m.recentJuneTurns(other).length, 1, "clearing is per user");
});

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-june-calendar-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
const Fastify = (await import("fastify")).default;
const juneRoutes = (await import("../routes/june.routes.js")).default;
const { db } = await import("../db/client.js");

const admin = { "x-user-role": "admin", "x-user-roles": "admin", "x-user-name": "jaco", "x-site-code": "main" };

test("June prepares an internal calendar change and only creates it after browser confirmation", async (t) => {
  t.after(async () => {
    db.close();
    rmSync(tempDir, { recursive: true, force: true });
  });
  const app = Fastify();
  await app.register(juneRoutes, { prefix: "/api/june" });
  t.after(() => app.close());

  const draft = await app.inject({
    method: "POST",
    url: "/api/june/gateway/execute",
    headers: admin,
    payload: {
      name: "june_prepare_internal_calendar_change",
      arguments: {
        action: "create",
        title: "Weekly maintenance forum",
        event_date: "2026-10-12",
        start_time: "08:00",
        end_time: "09:00",
        category: "meeting",
      },
    },
  });
  assert.equal(draft.statusCode, 200);
  const result = draft.json().result;
  assert.equal(result.state, "pending_confirmation");
  assert.equal(result.event.title, "Weekly maintenance forum");
  assert.match(result.approval.token, /^[0-9a-f-]{36}$/i);

  const before = await app.inject({ method: "GET", url: "/api/june/calendar/internal/events?days=30", headers: admin });
  assert.equal(before.statusCode, 200);
  assert.equal(before.json().calendar.events.length, 0);

  const otherAdmin = await app.inject({
    method: "POST",
    url: `/api/june/calendar/internal/approvals/${result.approval.token}/confirm`,
    headers: { ...admin, "x-user-name": "another-admin" },
    payload: {},
  });
  assert.equal(otherAdmin.statusCode, 404);

  const confirmed = await app.inject({
    method: "POST",
    url: `/api/june/calendar/internal/approvals/${result.approval.token}/confirm`,
    headers: admin,
    payload: {},
  });
  assert.equal(confirmed.statusCode, 200);
  assert.equal(confirmed.json().event.title, "Weekly maintenance forum");

  const after = await app.inject({ method: "GET", url: "/api/june/calendar/internal/events?days=30", headers: admin });
  assert.equal(after.statusCode, 200);
  assert.equal(after.json().calendar.events.length, 1);
  assert.equal(after.json().calendar.events[0].title, "Weekly maintenance forum");
});

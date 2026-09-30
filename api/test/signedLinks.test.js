import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { prestartPdfUrl, signRecordToken, verifyRecordToken } from "../utils/signedLinks.js";
import { isPublicAuthRequest } from "../auth/config.js";

test("a signed link opens only its own check and expires", () => {
  const prev = { a: process.env.IRONLOG_LINK_SECRET, b: process.env.IRONLOG_AUTH_SECRET };
  delete process.env.IRONLOG_LINK_SECRET;
  delete process.env.IRONLOG_AUTH_SECRET;
  try {
    const db = new Database(":memory:");
    const now = Date.parse("2026-09-30T08:00:00Z");
    const t = signRecordToken(db, "prestart-pdf", 42, { now });
    assert.equal(verifyRecordToken(db, "prestart-pdf", 42, t, { now }), true);
    assert.equal(verifyRecordToken(db, "prestart-pdf", 43, t, { now }), false, "another check number");
    assert.equal(verifyRecordToken(db, "other", 42, t, { now }), false, "another purpose");
    assert.equal(verifyRecordToken(db, "prestart-pdf", 42, t, { now: now + 15 * 86400000 }), false, "expired after 14 days");
    assert.equal(verifyRecordToken(db, "prestart-pdf", 42, `${t}x`, { now }), false);
    assert.equal(verifyRecordToken(db, "prestart-pdf", 42, "", { now }), false);
    const [exp] = t.split(".");
    assert.equal(verifyRecordToken(db, "prestart-pdf", 42, `${Number(exp) + 86400000}.${t.split(".")[1]}`, { now }), false, "stretched expiry");
    // The generated secret is kept, so links stay valid across restarts.
    assert.equal(verifyRecordToken(db, "prestart-pdf", 42, t, { now }), true);
    assert.match(prestartPdfUrl(db, 42), /^\/api\/reports\/prestart-check\/42\.pdf\?t=\d+\.[\w-]+$/);
    assert.equal(prestartPdfUrl(db, 0), null);
  } finally {
    if (prev.a !== undefined) process.env.IRONLOG_LINK_SECRET = prev.a;
    if (prev.b !== undefined) process.env.IRONLOG_AUTH_SECRET = prev.b;
  }
});

test("only the signed pre-start PDF path skips login", () => {
  assert.equal(isPublicAuthRequest("/api/reports/prestart-check/42.pdf", "GET"), true);
  assert.equal(isPublicAuthRequest("/api/reports/vehicle-ldv-check/42.pdf", "GET"), false);
  assert.equal(isPublicAuthRequest("/api/reports/prestart-check/42.pdf", "POST"), false);
  assert.equal(isPublicAuthRequest("/api/reports/vehicle-ldv-checks.pdf", "GET"), false);
});

import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";

const src = fs.readFileSync(new URL("../../web/app/developer.js", import.meta.url), "utf8");
const html = fs.readFileSync(new URL("../../web/index.html", import.meta.url), "utf8");
const css = fs.readFileSync(new URL("../../web/styles.css", import.meta.url), "utf8");

function el(hidden) {
  return { hidden, innerHTML: "", attrs: {}, setAttribute(k, v) { this.attrs[k] = v; } };
}

// Loads developer.js with a tiny fake page and a canned /status answer.
function setup(status) {
  const els = { devSetup: el(true), devReady: el(true), devChecklist: el(false) };
  const ctx = {
    qs: (id) => els[id] || null,
    escapeHtml: (s) => String(s),
    fetchJson: async () => status,
    API: "",
    window: {},
    document: {},
    console,
  };
  vm.createContext(ctx);
  vm.runInContext(src, ctx);
  return { els, ctx };
}

const ready = { github_configured: true, repo: "a/b", checks: { repo: { ok: true }, workflow: { ok: true } } };

test("developer setup: ready shows one green line and hides the orange box", async () => {
  const { els, ctx } = setup(ready);
  await ctx.loadDeveloperSetup();
  assert.equal(els.devReady.hidden, false);
  assert.equal(els.devSetup.hidden, true);
  assert.equal(els.devReady.attrs["aria-expanded"], "false");
});

test("developer setup: clicking the green line opens the checklist and clicking again closes it", async () => {
  const { els, ctx } = setup(ready);
  await ctx.loadDeveloperSetup();
  ctx.toggleDeveloperSetup();
  assert.equal(els.devSetup.hidden, false);
  assert.equal(els.devReady.attrs["aria-expanded"], "true");
  ctx.toggleDeveloperSetup();
  assert.equal(els.devSetup.hidden, true);
});

test("developer setup: missing workflow keeps the full orange box and no green line", async () => {
  const { els, ctx } = setup({ ...ready, checks: { repo: { ok: true }, workflow: { ok: false, error: "missing" } } });
  await ctx.loadDeveloperSetup();
  assert.equal(els.devReady.hidden, true);
  assert.equal(els.devSetup.hidden, false);
});

test("developer setup: GitHub not connected keeps the full orange box", async () => {
  const { els, ctx } = setup({ github_configured: false });
  await ctx.loadDeveloperSetup();
  assert.equal(els.devReady.hidden, true);
  assert.equal(els.devSetup.hidden, false);
});

test("developer setup: green line is in the page and styled with theme tokens", () => {
  assert.match(html, /id="devReady"[^>]*>✓ Ready: Claude is connected\.</);
  const rule = css.split("\n").find((l) => l.startsWith("#tab-developer .dev-ready {"));
  assert.ok(rule, "dev-ready style is missing");
  assert.match(rule, /var\(--success\)/);
  assert.doesNotMatch(rule, /#[0-9a-f]{3,6}\b/i, "use colour tokens, not raw hex");
});

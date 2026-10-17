import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

// Ease of use, part 1: one menu on every desk page (web/app-nav.js).
const web = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "web");
const ctx = { window: {}, localStorage: { getItem: () => null } };
vm.runInNewContext(fs.readFileSync(path.join(web, "app-nav.js"), "utf8"), ctx);
const nav = ctx.window.IronlogNav;

test("every menu item opens something real", () => {
  const index = fs.readFileSync(path.join(web, "index.html"), "utf8");
  for (const sec of nav.NAV_SECTIONS) {
    for (const it of sec.items) {
      if (it.tab) assert.ok(index.includes(`id="tab-${it.tab}"`), `${it.label}: no screen tab-${it.tab} in index.html`);
      if (it.href) assert.ok(fs.existsSync(path.join(web, it.href)), `${it.label}: ${it.href} does not exist`);
    }
  }
  const tabs = nav.NAV_SECTIONS.flatMap((s) => s.items.map((i) => i.tab).filter(Boolean));
  for (const t of nav.PRODUCTION_SITE_TABS.filter((x) => x !== "maintenance")) assert.ok(tabs.includes(t), `screen ${t} is missing from the menu`);
});

test("the desk pages all use the shared menu", () => {
  for (const [page, call] of [["index.html", 'data-ironlog-nav="tabs"'], ["workorders.html", "mountPage"], ["breakdown-ops.html", "mountPage"], ["api-docs.html", "mountPage"], ["maintenance.html", "mountIntoSidebar"]]) {
    const html = fs.readFileSync(path.join(web, page), "utf8");
    assert.match(html, /<script src="\.\/app-nav\.js\?v=\d+"><\/script>/, `${page} loads app-nav.js`);
    assert.ok(html.includes(call), `${page} shows the shared menu`);
  }
  assert.equal((fs.readFileSync(path.join(web, "index.html"), "utf8").match(/class="nav-item"/g) || []).length, 0, "no second hand-written menu in index.html");
});

test("each role sees the same screens as before, and the new links follow the right rules", () => {
  assert.deepEqual([...nav.effectiveTabs({ roles: ["artisan"] })], ["mywork", "workshop", "maintenance", "vehicle", "tasks", "ironmind"]);
  const stores = new Set(nav.effectiveTabs({ roles: ["stores"] }));
  assert.ok(stores.has("stock") && stores.has("procurement") && !stores.has("finance") && !stores.has("admin"));
  const admin = new Set(nav.effectiveTabs({ roles: ["admin"] }));
  assert.ok(admin.has("admin") && admin.has("developer") && admin.has("uploads"));
  assert.equal(new Set(nav.effectiveTabs({ roles: ["admin"], unlicensed: new Set(["fuel"]) })).has("fuel"), false, "unlicensed screens are hidden");
  assert.deepEqual([...nav.effectiveTabs({ roles: ["operator"], override: ["fuel"] })], ["mywork", "fuel", "ironmind", "tasks"], "an admin's per-user list wins");

  const item = (key) => nav.NAV_SECTIONS.flatMap((s) => s.items).find((i) => i.key === key || i.tab === key);
  assert.equal(nav.itemVisible(item("workorders"), stores, ["stores"]), false, "Work Orders goes with Maintenance");
  assert.equal(nav.itemVisible(item("workorders"), new Set(nav.effectiveTabs({ roles: ["artisan"] })), ["artisan"]), true);
  assert.equal(nav.itemVisible(item("api-docs"), admin, ["supervisor"]), false, "API documentation is for administrators");
  assert.equal(nav.itemVisible(item("tech-portal"), stores, ["artisan"]), true);
  assert.match(nav.sectionsHtml({ mode: "links" }), /href="index\.html\?tab=stock"/, "other pages link back to the main screen");
  assert.match(nav.sectionsHtml({ mode: "tabs" }), /href="#"[^>]*data-tab="stock"/, "the main screen switches in place");
});

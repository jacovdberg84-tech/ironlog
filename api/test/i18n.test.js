import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import { fileURLToPath } from "node:url";

// Ease of use: Portuguese on every desk page (web/i18n.js + web/i18n-pt.js).
const web = path.join(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "web");
const ctx = { window: {}, localStorage: { getItem: () => null } };
ctx.globalThis = ctx.window;
vm.createContext(ctx);
for (const f of ["i18n.js", "i18n-pt.js", "app-nav.js"]) vm.runInContext(fs.readFileSync(path.join(web, f), "utf8").replace(/^IronlogI18n\./gm, "window.IronlogI18n."), ctx);
const { IronlogI18n: i18n, IronlogNav: nav } = ctx.window;
const pt = (s) => i18n.t(s, "pt");
const known = (s) => i18n.has(s, "pt");
const decode = (s) => s.replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&apos;|&#39;/g, "'").replace(/&quot;/g, '"').replace(/&lt;/g, "<").replace(/&gt;/g, ">");

test("phrases, patterns and what is left alone", () => {
  assert.equal(pt("Work Orders"), "Ordens de trabalho");
  assert.equal(pt("  Log a breakdown "), "Registar uma avaria", "spaces around a phrase do not matter");
  assert.equal(pt("Hi sipho, here's what needs attention"), "Olá sipho, isto precisa da sua atenção");
  assert.equal(pt("3 machines down"), "3 máquinas paradas");
  assert.equal(pt("Purchase request"), "Purchase request", "numbers-only placeholders never catch typed words");
  assert.equal(pt("A302AM"), "A302AM", "machine codes and anything unknown stay as they are");
  assert.equal(i18n.t("Work Orders", "en"), "Work Orders");
  // Costing gaps come from the server with machine and part codes inside; the codes stay.
  assert.equal(pt("WO #12 (DT03) has no labour hours"), "OT #12 (DT03) não tem horas de mão de obra");
  assert.equal(pt("2 parts at $0"), "2 peças a $0");
});

test("the menu and the daily-use screens have Portuguese for every phrase", () => {
  const missing = [];
  for (const sec of nav.NAV_SECTIONS) {
    for (const s of [sec.title, ...sec.items.flatMap((i) => [i.label, i.title])]) if (!known(s)) missing.push(`menu: ${s}`);
  }
  // Visible text in the screens technicians, operators and foremen use most.
  const index = fs.readFileSync(path.join(web, "index.html"), "utf8");
  const sectionOf = (id) => {
    const start = index.indexOf(`id="${id}"`);
    const end = index.indexOf('\n    <section class="panel', start + 10);
    return index.slice(start, end > 0 ? end : undefined);
  };
  const pages = { ...Object.fromEntries(["tab-mywork", "tab-daily", "tab-fuel", "tab-lube", "tab-vehicle", "tab-workshop", "tab-tasks"].map((id) => [id, sectionOf(id)])),
    "workorders.html": fs.readFileSync(path.join(web, "workorders.html"), "utf8"), "breakdown-ops.html": fs.readFileSync(path.join(web, "breakdown-ops.html"), "utf8") };
  for (const [where, html] of Object.entries(pages)) {
    const text = html.replace(/<(script|style|code|pre|svg|textarea)\b[\s\S]*?<\/\1>/gi, " ");
    for (const m of text.matchAll(/>([^<>]+)</g)) {
      const s = decode(m[1]).replace(/\s+/g, " ").trim();
      if (!/[A-Za-z]{3}/.test(s) || /[ãõçáéíóúâêô]/.test(s) || /^[A-Z0-9_./-]+$/.test(s) || /^(e\.g\.|https?:)/.test(s) || /^\(.*\)$/.test(s)) continue;
      if (/^(IRONLOG|IronLog|Borris|June|FAMS|Cartrack|Komatsu|Bell|Sergio|Microsoft Outlook|API base URL|Emma Tool Gateway|ONLINE|km|L\/hr|MTD|YTD|PT ⇄ EN|FSC Pilot|CSV format:)/.test(s)) continue;
      if (!known(s)) missing.push(`${where}: ${s}`);
    }
  }
  assert.deepEqual(missing, [], `Add these to web/i18n-pt.js:\n${missing.join("\n")}`);
});

test("every desk page loads the translator before the menu", () => {
  for (const page of ["index.html", "workorders.html", "breakdown-ops.html", "api-docs.html", "maintenance.html"]) {
    const html = fs.readFileSync(path.join(web, page), "utf8");
    const a = html.indexOf("./i18n.js?v="), b = html.indexOf("./i18n-pt.js?v="), c = html.indexOf("./app-nav.js?v=");
    assert.ok(a > 0 && a < b && b < c, `${page}: i18n.js, then i18n-pt.js, then app-nav.js`);
  }
  assert.match(nav.sectionsHtml({ mode: "links" }), /data-i18n-switch="pt"/, "the EN | PT switch is in the menu");
});

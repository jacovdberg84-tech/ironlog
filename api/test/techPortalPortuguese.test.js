import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const web = new URL("../../web/", import.meta.url);
const read = (f) => readFileSync(new URL(f, web), "utf8");

function portugueseText() {
  const sandbox = { window: {} };
  vm.runInNewContext(read("tech-portal-pt.js"), sandbox);
  return sandbox.window.TechPortalPT;
}

/** Double-quoted literals inside a call like T( ... ), skipping values compared with ===. */
function literalsInCalls(src, fn) {
  const out = new Set();
  const re = new RegExp(`(?<![\\w.])${fn}\\(`, "g");
  let m;
  while ((m = re.exec(src))) {
    let i = m.index + m[0].length;
    let depth = 1;
    let quote = null;
    const start = i;
    for (; i < src.length && depth; i += 1) {
      const c = src[i];
      if (quote) {
        if (c === "\\") i += 1;
        else if (c === quote) quote = null;
        continue;
      }
      if (c === '"' || c === "'") quote = c;
      else if (c === "(") depth += 1;
      else if (c === ")") depth -= 1;
    }
    const span = src.slice(start, i - 1);
    for (const lit of span.matchAll(/(===\s*)?"((?:[^"\\]|\\.)*)"/g)) {
      if (!lit[1] && lit[2]) out.add(JSON.parse(`"${lit[2]}"`));
    }
  }
  return out;
}

/** Literals in a table the code passes through T() (states, statuses, button labels). */
function literalsInBlock(src, marker, end) {
  const i = src.indexOf(marker);
  assert.ok(i >= 0, `missing ${marker}`);
  const block = src.slice(i, src.indexOf(end, i));
  const out = new Set();
  for (const lit of block.matchAll(/"((?:[^"\\]|\\.)*)"/g)) {
    if (!/^[a-z_]+$/.test(lit[1])) out.add(JSON.parse(`"${lit[1]}"`));
  }
  return out;
}

test("every portal screen text has a Portuguese translation", () => {
  const PT = portugueseText();
  const portal = read("tech-portal.js");
  const terminal = read("technician-terminal.js");
  const html = read("technician-terminal.html");
  const needed = new Set([
    ...literalsInCalls(portal, "T"),
    ...literalsInCalls(portal, "N"),
    ...literalsInCalls(terminal, "tt"),
    ...literalsInBlock(portal, "const STATE_TEXT", "};"),
    ...literalsInBlock(portal, "const WO_STATUS", "};"),
    ...literalsInBlock(portal, "const reqStatus", "};"),
    ...literalsInBlock(portal, "const ACTIONS", "done: []"),
    ...literalsInBlock(portal, "const SHIFT_FIELDS", "];"),
    ...literalsInBlock(portal, "const ideas", ".map("),
    ...[...html.matchAll(/data-t="([^"]+)"/g)].map((m) => m[1].replace(/&amp;/g, "&")),
  ]);
  // Not screen text: button style classes and punctuation.
  const notText = (k) => !/[A-Za-z]/.test(k) || /^(primary|ok|warn|big)( (primary|ok|warn|big))*$/.test(k);
  const missing = [...needed].filter((k) => !notText(k) && !PT[k]);
  assert.deepEqual(missing, [], `missing Portuguese for: ${missing.join(" | ")}`);
  assert.ok(needed.size > 200, `found ${needed.size} texts`);
});

test("Portuguese keeps every {placeholder} of the English text", () => {
  const PT = portugueseText();
  const holes = (s) => [...String(s).matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort().join(",");
  const wrong = Object.entries(PT).filter(([en, pt]) => holes(en) !== holes(pt)).map(([en]) => en);
  assert.deepEqual(wrong, []);
});

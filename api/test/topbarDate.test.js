import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";

const css = fs.readFileSync(new URL("../../web/styles.css", import.meta.url), "utf8");
const html = fs.readFileSync(new URL("../../web/index.html", import.meta.url), "utf8");

test("top bar date box is wide enough to show the whole date", () => {
  const rule = css.match(/\.icon-btn-label input\[type="date"\]\s*\{([^}]*)\}/);
  assert.ok(rule, "date box needs its own width rule");
  const min = rule[1].match(/min-width:\s*(\d+)px/);
  assert.ok(min && Number(min[1]) >= 110, "date box must be at least 110px wide");
});

test("index.html loads a fresh styles.css version", () => {
  assert.match(html, /styles\.css\?v=date-card-1/);
});

import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";
import { pageScriptFiles, readPageSource } from "./webAppSource.js";

const pages = [
  { page: "index.html", dir: "app", first: "app/core.js" },
  { page: "maintenance.html", dir: "maintenance", first: "maintenance/core.js" },
];

for (const { page, dir, first } of pages) {
  test(`${page} loads every web/${dir} file exactly once`, () => {
    const listed = pageScriptFiles(page, dir);
    const onDisk = fs.readdirSync(new URL(`../../web/${dir}/`, import.meta.url)).filter((f) => f.endsWith(".js")).map((f) => `${dir}/${f}`);
    assert.equal(new Set(listed).size, listed.length, "a file is listed twice");
    assert.deepEqual([...listed].sort(), onDisk.sort());
    assert.equal(listed[0], first, "core.js defines shared helpers and must load first");
  });

  test(`each web/${dir} file parses as a classic script`, () => {
    for (const f of pageScriptFiles(page, dir)) {
      assert.doesNotThrow(() => new vm.Script(fs.readFileSync(new URL(`../../web/${f}`, import.meta.url), "utf8"), { filename: f }), f);
    }
    assert.doesNotThrow(() => new vm.Script(readPageSource(page, dir)));
  });
}

import test from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import vm from "node:vm";
import { webAppFiles, readWebAppSource } from "./webAppSource.js";

test("index.html loads every web/app file exactly once", () => {
  const listed = webAppFiles();
  const onDisk = fs.readdirSync(new URL("../../web/app/", import.meta.url)).filter((f) => f.endsWith(".js")).map((f) => `app/${f}`);
  assert.equal(new Set(listed).size, listed.length, "a file is listed twice");
  assert.deepEqual([...listed].sort(), onDisk.sort());
  assert.equal(listed[0], "app/core.js", "core.js defines shared helpers and must load first");
});

test("each web/app file parses as a classic script", () => {
  for (const f of webAppFiles()) {
    assert.doesNotThrow(() => new vm.Script(fs.readFileSync(new URL(`../../web/${f}`, import.meta.url), "utf8"), { filename: f }), f);
  }
  assert.doesNotThrow(() => new vm.Script(readWebAppSource()));
});

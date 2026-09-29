import test from "node:test";
import assert from "node:assert/strict";
import { getMachinePrestartTemplate, listMachinePrestartProfiles } from "../utils/machinePrestartTemplates.js";
import { PRESTART_PT, withPortuguese } from "../utils/prestartPortuguese.js";

test("every machine pre-start title, section and check has Portuguese", () => {
  const missing = [];
  for (const { id } of listMachinePrestartProfiles()) {
    const t = getMachinePrestartTemplate(id);
    for (const text of [t.title, ...t.sections.flatMap((s) => [s.title, ...s.items.map((i) => i.label)])]) {
      if (!PRESTART_PT[text]) missing.push(`${id}: ${text}`);
    }
  }
  assert.deepEqual(missing, []);
});

test("withPortuguese keeps English and adds label_pt", () => {
  const t = withPortuguese(getMachinePrestartTemplate("tipper_truck"));
  assert.equal(t.title, "Tipper truck pre-start");
  assert.equal(t.title_pt, "Pré-arranque do camião basculante");
  const first = t.sections[0].items[0];
  assert.equal(first.label, "Engine oil OK");
  assert.equal(first.label_pt, "Óleo do motor OK");
});

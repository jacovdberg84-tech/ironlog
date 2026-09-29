import test from "node:test";
import assert from "node:assert/strict";
import Database from "better-sqlite3";
import { cleanTranslation, translateText, translationMessages } from "../utils/translate.js";
import { addEnglishToFaultComments } from "../utils/prestartFaults.js";

test("prompt names the target language and keeps the user's text as is", () => {
  const m = translationMessages("Travão fraco", "en");
  assert.match(m[0].content, /into English/);
  assert.equal(m[1].content, "Travão fraco");
  assert.match(translationMessages("x", "pt")[0].content, /Portuguese as spoken in Mozambique/);
});

test("model replies are cleaned of quotes and prefixes", () => {
  assert.equal(cleanTranslation('"Weak brake"'), "Weak brake");
  assert.equal(cleanTranslation("Translation: Weak brake"), "Weak brake");
  assert.equal(cleanTranslation("  Weak brake \n"), "Weak brake");
});

test("translateText returns null without a chat function or on failure", async () => {
  assert.equal(await translateText("olá"), null);
  assert.equal(await translateText("olá", { chat: async () => null }), null);
  assert.equal(await translateText("  ", { chat: async () => "x" }), "");
  assert.equal(await translateText("pneu furado", { chat: async () => "Flat tyre" }), "Flat tyre");
});

test("Portuguese fault comments get an English copy on the work order once", async () => {
  const db = new Database(":memory:");
  db.exec(`CREATE TABLE work_orders (id INTEGER PRIMARY KEY, job_description TEXT)`);
  db.prepare("INSERT INTO work_orders VALUES (1, ?)").run("Pre-start faults reported by Ana on 2026-09-29:\n- Brakes: pedal mole\n- Tyres");
  const faults = [{ key: "b", label: "Brakes", comment: "pedal mole" }, { key: "t", label: "Tyres", comment: "" }];
  let calls = 0;
  const translate = async () => { calls += 1; return "soft pedal"; };
  assert.equal(await addEnglishToFaultComments(db, 1, faults, translate), true);
  assert.equal(db.prepare("SELECT job_description d FROM work_orders").get().d,
    "Pre-start faults reported by Ana on 2026-09-29:\n- Brakes: pedal mole (EN: soft pedal)\n- Tyres");
  assert.equal(await addEnglishToFaultComments(db, 1, faults, translate), false);
  assert.equal(calls, 1);
  assert.equal(await addEnglishToFaultComments(db, 1, faults, async () => "pedal mole"), false, "same text is not repeated");
});

import test from "node:test";
import assert from "node:assert/strict";
import { createBorrisQueue } from "../utils/borrisQueue.js";

test("jobs run one at a time, in order, and keep their result", async () => {
  let active = 0;
  let maxActive = 0;
  const order = [];
  const q = createBorrisQueue({
    run: async (key) => {
      active += 1; maxActive = Math.max(maxActive, active); order.push(key);
      await new Promise((r) => setTimeout(r, 10));
      active -= 1;
      if (key === "bad") throw new Error("model offline");
      return { key };
    },
  });
  assert.equal(q.enqueue("a").status, "working");
  assert.equal(q.enqueue("b").status, "queued");
  assert.equal(q.enqueue("b").status, "queued", "same key is not queued twice");
  assert.equal(q.enqueue("bad").status, "queued");
  await q.idle();
  assert.deepEqual(order, ["a", "b", "bad"]);
  assert.equal(maxActive, 1);
  assert.deepEqual(q.status("a").result, { key: "a" });
  assert.equal(q.status("bad").status, "failed");
  assert.equal(q.status("bad").error, "model offline");
  assert.equal(q.status("nope").status, "none");
});

test("a finished job is reused unless refreshed; too many waiting is refused", async () => {
  let runs = 0;
  const q = createBorrisQueue({ run: async () => { runs += 1; return runs; }, maxWaiting: 1 });
  q.enqueue("a");
  await q.idle();
  assert.equal(q.enqueue("a").result, 1);
  q.enqueue("a", { refresh: true });
  await q.idle();
  assert.equal(q.status("a").result, 2);
  const slow = createBorrisQueue({ run: () => new Promise((r) => setTimeout(r, 30)), maxWaiting: 1 });
  slow.enqueue("x");
  slow.enqueue("y");
  assert.equal(slow.enqueue("z").status, "busy");
  await slow.idle();
});

test("old results are pruned after keepMs", async () => {
  let t = 1000;
  const q = createBorrisQueue({ run: async () => "ok", keepMs: 100, now: () => t });
  q.enqueue("a");
  await q.idle();
  t += 500;
  q.enqueue("b");
  assert.equal(q.status("a").status, "none");
  await q.idle();
});

import test from "node:test";
import assert from "node:assert/strict";
import { createRequire } from "node:module";

const { addRoute } = createRequire(import.meta.url)("../../deploy/tunnel-route.cjs");

const config = `tunnel: 6f1c0e2a-1111-2222-3333-444455556666
credentials-file: /root/.cloudflared/6f1c0e2a-1111-2222-3333-444455556666.json

ingress:
  - hostname: ironlog.ironlogafrica.com
    service: http://localhost:3001
  - service: http_status:404
`;

test("the pilot route goes just above the catch-all; AML's route is untouched", () => {
  const out = addRoute(config, "pilot.ironlogafrica.com", "3101");
  assert.equal(out.result, "added");
  assert.equal(out.text, config.replace("  - service: http_status:404", "  - hostname: pilot.ironlogafrica.com\n    service: http://localhost:3101\n  - service: http_status:404"));
  assert.equal(addRoute(out.text, "pilot.ironlogafrica.com", "3101").result, "present", "running it twice changes nothing");
  assert.equal(addRoute(out.text, "pilot2.ironlogafrica.com", "3102").result, "added", "a similar name is not confused");
});

test("a config it does not understand is refused, not guessed at", () => {
  assert.throws(() => addRoute("tunnel: x\n", "pilot.ironlogafrica.com", "3101"), /no ingress/);
  assert.throws(() => addRoute("ingress:\n  - hostname: a.b\n    service: http://localhost:3001\n", "pilot.ironlogafrica.com", "3101"), /no catch-all/);
});

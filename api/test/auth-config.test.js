import test from "node:test";
import assert from "node:assert/strict";
import { isAuthRequired } from "../auth/config.js";

test("requires authentication for the default network host", () => {
  assert.equal(isAuthRequired({}), true);
  assert.equal(isAuthRequired({ HOST: "0.0.0.0" }), true);
});

test("allows local-first desktop and loopback runtimes", () => {
  assert.equal(isAuthRequired({ IRONLOG_DESKTOP: "1" }), false);
  assert.equal(isAuthRequired({ HOST: "127.0.0.1" }), false);
  assert.equal(isAuthRequired({ HOST: "localhost" }), false);
  assert.equal(isAuthRequired({ HOST: "::1" }), false);
});

test("honors an explicit authentication setting", () => {
  assert.equal(isAuthRequired({ IRONLOG_AUTH_REQUIRED: "true", HOST: "localhost" }), true);
  assert.equal(isAuthRequired({ IRONLOG_AUTH_REQUIRED: "1", HOST: "localhost" }), true);
  assert.equal(isAuthRequired({ IRONLOG_AUTH_REQUIRED: "false", HOST: "0.0.0.0" }), false);
  assert.equal(isAuthRequired({ IRONLOG_AUTH_REQUIRED: "0", HOST: "0.0.0.0" }), false);
});

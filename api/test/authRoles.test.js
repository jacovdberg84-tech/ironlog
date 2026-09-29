import test from "node:test";
import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import os from "node:os";
import path from "node:path";

const tempDir = mkdtempSync(path.join(os.tmpdir(), "ironlog-auth-roles-"));
process.env.DB_PATH = path.join(tempDir, "ironlog.db");
process.env.IRONLOG_AUTH_REQUIRED = "0";
const { ironlogAuthHook, primaryRouteRole } = await import("../auth/hook.js");

async function rolesFor(roles, url) {
  const req = { url, method: "POST", headers: { "x-user-role": roles[0], "x-user-roles": roles.join(",") } };
  await ironlogAuthHook(req, { code: () => ({ send: () => {} }) });
  return { primary: req.headers["x-user-role"], all: req.headers["x-user-roles"].split(",") };
}

test("storemen pass the 'stores' checks used by stock, work order and purchasing routes", async (t) => {
  t.after(() => rmSync(tempDir, { recursive: true, force: true }));
  for (const url of ["/api/stock/lube-issue", "/api/workorders/12/issue", "/api/procurement/requisitions"]) {
    const r = await rolesFor(["storeman"], url);
    assert.equal(r.primary, "stores", url);
    assert.ok(r.all.includes("storeman") && r.all.includes("stores"), url);
  }
});

test("primary role keeps admins first and maps the newer role names", () => {
  assert.equal(primaryRouteRole(["storeman", "admin", "stores"]), "admin");
  assert.equal(primaryRouteRole(["workshop_admin", "supervisor"]), "supervisor");
  assert.equal(primaryRouteRole(["plant_clerk", "operator"]), "operator");
  assert.equal(primaryRouteRole(["plant_manager"]), "plant_manager");
});

test("role checks accept any role the user holds", async () => {
  const { holdsAnyRole } = await import("../utils/request.js");
  const req = (role, roles) => ({ headers: { ...(role ? { "x-user-role": role } : {}), ...(roles ? { "x-user-roles": roles } : {}) } });
  assert.equal(holdsAnyRole(req("plant_manager", "plant_manager,storeman,stores"), ["admin", "supervisor", "stores"]), true);
  assert.equal(holdsAnyRole(req("plant_manager", "plant_manager"), ["admin", "supervisor", "stores"]), false);
  assert.equal(holdsAnyRole(req(null, null), ["admin"]), true, "no role headers keeps the old admin default");
});

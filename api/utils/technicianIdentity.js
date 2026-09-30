// IRONLOG/api/utils/technicianIdentity.js — is this work order the signed-in
// technician's? Work orders store the assignee as a login username or, on older
// rows, a display name, so both are matched against the users table.

function hasUsersTable(db) {
  return Boolean(db.prepare(`SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'users'`).get());
}

/** Lower-case username and full name that identify one person. */
export function technicianIdentityKeys(db, nameOrUsername) {
  const raw = String(nameOrUsername || "").trim().toLowerCase();
  const keys = new Set();
  if (!raw) return keys;
  keys.add(raw);
  if (!hasUsersTable(db)) return keys;
  const byUser = db.prepare(`
    SELECT username, full_name FROM users
    WHERE LOWER(TRIM(username)) = ?
    LIMIT 1
  `).get(raw);
  if (byUser) {
    keys.add(String(byUser.username || "").trim().toLowerCase());
    const fn = String(byUser.full_name || "").trim().toLowerCase();
    if (fn) keys.add(fn);
  }
  const byName = db.prepare(`
    SELECT username, full_name FROM users
    WHERE LOWER(TRIM(COALESCE(full_name, ''))) = ?
    LIMIT 1
  `).get(raw);
  if (byName) {
    keys.add(String(byName.username || "").trim().toLowerCase());
    keys.add(raw);
  }
  return keys;
}

/** Match assigned technician to logged-in user (username or legacy display name). */
export function technicianMatchesUser(db, assignedName, userName) {
  const assignedKeys = technicianIdentityKeys(db, assignedName);
  const userKeys = technicianIdentityKeys(db, userName);
  if (!assignedKeys.size || !userKeys.size) return false;
  for (const k of userKeys) {
    if (assignedKeys.has(k)) return true;
  }
  return false;
}

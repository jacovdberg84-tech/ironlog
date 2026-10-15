// IRONLOG/api/db/client.js (SQLite)
import Database from "better-sqlite3";
import path from "node:path";
// Settings first (api/.env, or an instance's own file): migrate imports this early.
import "../loadEnv.js";

const dbPath = process.env.DB_PATH || "./db/ironlog.db";
export const dbPathResolved = path.resolve(path.normalize(dbPath));

// Windows-safe + consistent
export const db = new Database(dbPathResolved);

// Enforce FK rules (SQLite default is OFF unless enabled)
db.pragma("foreign_keys = ON");
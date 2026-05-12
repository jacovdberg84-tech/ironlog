// IRONLOG/api/db/client.js (SQLite)
import Database from "better-sqlite3";
import dotenv from "dotenv";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirnameDb = path.dirname(fileURLToPath(import.meta.url));
// Load api/.env even when Node cwd is the repo root (migrate imports this early).
dotenv.config({ path: path.join(__dirnameDb, "..", ".env") });
dotenv.config();

const dbPath = process.env.DB_PATH || "./db/ironlog.db";
export const dbPathResolved = path.resolve(path.normalize(dbPath));

// Windows-safe + consistent
export const db = new Database(dbPathResolved);

// Enforce FK rules (SQLite default is OFF unless enabled)
db.pragma("foreign_keys = ON");
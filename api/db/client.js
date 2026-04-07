// IRONLOG/api/db/client.js (SQLite)
import Database from "better-sqlite3";
import dotenv from "dotenv";
import path from "node:path";

dotenv.config();

const dbPath = process.env.DB_PATH || "./db/ironlog.db";
export const dbPathResolved = path.resolve(path.normalize(dbPath));

// Windows-safe + consistent
export const db = new Database(dbPathResolved);

// Enforce FK rules (SQLite default is OFF unless enabled)
db.pragma("foreign_keys = ON");
// IRONLOG/api/db/client.js (SQLite)
import Database from "better-sqlite3";
import dotenv from "dotenv";
import path from "node:path";

dotenv.config();

const dbPath = process.env.DB_PATH || "./db/ironlog.db";

// Windows-safe + consistent
export const db = new Database(path.normalize(dbPath));

// Enforce FK rules (SQLite default is OFF unless enabled)
db.pragma("foreign_keys = ON");
// IRONLOG/api/db/migrate.js (SQLite)
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { db } from "./client.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function migrate() {
  const schemaPath = path.join(__dirname, "schema.sql");

  if (!fs.existsSync(schemaPath)) {
    throw new Error(`schema.sql not found at: ${schemaPath}`);
  }

  const sql = fs.readFileSync(schemaPath, "utf8");

  console.log("🧱 Running IRONLOG migration (SQLite schema.sql)...");
  const tx = db.transaction(() => {
    db.exec(sql);
  });

  tx();
  console.log("✅ Migration complete.");
}

try {
  migrate();
} catch (err) {
  console.error("❌ Migration failed:", err.message);
  process.exitCode = 1;
}
// IRONLOG/api/scripts/first-admin.js — adds the first administrator to a new company instance.
// Usage (run by deploy/ironlog-instance.sh):
//   IRONLOG_ENV_FILE=/data/ironlog-instances/<name>/instance.env node scripts/first-admin.js <username> "<Full name>"
// Prints a one-time setup code; the person signs in with it and chooses a password.
import "../loadEnv.js";
import Database from "better-sqlite3";
import { createFirstAdmin } from "../utils/firstAdmin.js";

if (!process.env.IRONLOG_ENV_FILE || !process.env.DB_PATH) {
  console.error("Only for company instances: set IRONLOG_ENV_FILE to the instance's settings file.");
  process.exit(1);
}
const db = new Database(process.env.DB_PATH, { fileMustExist: true });
const result = createFirstAdmin(db, { username: process.argv[2], full_name: process.argv[3] });
db.close();
if (!result.ok) {
  console.error(result.error);
  process.exit(result.error.startsWith("This instance already") ? 0 : 1);
}
console.log(`First administrator: ${result.username}`);
console.log(`One-time setup code: ${result.setup_code} (valid until ${result.expires_at} UTC)`);
console.log("On the sign-in page enter the username, put this code in 'Setup code (first-time only)' and type a new password.");

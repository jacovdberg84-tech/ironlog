const crypto = require("crypto");
const Database = require("better-sqlite3");

const db = new Database("./db/ironlog.db");
const pwd = "ChangeMe123!";

const salt = crypto.randomBytes(16);
const hash = crypto.scryptSync(pwd, salt, 64);
const stored = "scrypt$" + salt.toString("base64") + "$" + hash.toString("base64");

db.prepare("UPDATE users SET password_hash=? WHERE username=?").run(stored, "admin");
console.log("Admin password set. Username: admin");

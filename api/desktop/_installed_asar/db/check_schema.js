import { db } from "./client.js";

console.log("breakdowns columns:");
console.table(db.prepare("PRAGMA table_info(breakdowns)").all());

console.log("\nHas breakdown_downtime_logs table?");
console.log(db.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name='breakdown_downtime_logs'").get());
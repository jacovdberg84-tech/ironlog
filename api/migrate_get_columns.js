import { db } from "./db/client.js";

try {

  db.exec(`
    ALTER TABLE breakdowns ADD COLUMN get_used INTEGER NOT NULL DEFAULT 0;
  `);

} catch(e) {}

try {

  db.exec(`
    ALTER TABLE breakdowns ADD COLUMN get_hours_fitted REAL;
  `);

} catch(e) {}

try {

  db.exec(`
    ALTER TABLE breakdowns ADD COLUMN get_hours_changed REAL;
  `);

} catch(e) {}

console.log("GET columns migration complete");
process.exit();
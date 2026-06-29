#!/usr/bin/env node
/**
 * Test IRONLOG SMTP (nodemailer) configuration.
 *
 * Usage:
 *   node scripts/test-smtp.js you@company.com
 *   node scripts/test-smtp.js --bootstrap you@company.com
 */

import path from "path";
import { fileURLToPath } from "url";
import dotenv from "dotenv";
import {
  bootstrapSmtpFromEnv,
  getSmtpSettingsRow,
  sendIronlogMail,
  smtpPublicPayload,
} from "../utils/mail.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
dotenv.config({ path: path.join(__dirname, "../.env") });
dotenv.config({ path: path.join(process.cwd(), ".env") });

async function main() {
  const args = process.argv.slice(2).filter((a) => a !== "--bootstrap" && a !== "--force");
  const doBootstrap = process.argv.includes("--bootstrap") || process.argv.includes("--force");
  const to = args[0] || String(process.env.SMTP_TEST_TO || "").trim();

  if (doBootstrap) {
    const boot = bootstrapSmtpFromEnv({
      force: process.argv.includes("--force"),
      log: console,
    });
    console.log("[smtp] bootstrap:", boot);
  }

  const row = getSmtpSettingsRow();
  console.log("[smtp] settings:", smtpPublicPayload(row));

  if (!to) {
    console.log("\nUsage: node scripts/test-smtp.js [--bootstrap] recipient@company.com");
    console.log("Or set SMTP_TEST_TO in api/.env");
    process.exitCode = 1;
    return;
  }

  const result = await sendIronlogMail({
    to,
    subject: "IRONLOG SMTP test email",
    text: `SMTP test successful at ${new Date().toISOString()}\n\nIf you received this, nodemailer is configured correctly.`,
  });

  if (!result.ok) {
    console.error("[smtp] send failed:", result.error);
    process.exitCode = 1;
    return;
  }

  console.log(`[smtp] test email sent to ${result.recipients.join(", ")}`);
  if (result.messageId) console.log("[smtp] message id:", result.messageId);
}

main().catch((err) => {
  console.error("[smtp] test failed:", err?.message || err);
  process.exitCode = 1;
});

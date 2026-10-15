// deploy/tunnel-route.cjs — adds one hostname rule to a cloudflared config, above the catch-all.
// Usage: node tunnel-route.cjs <config file> <hostname> <port>   → prints "added" or "present".
const fs = require("fs");

function addRoute(text, host, port) {
  const lines = text.split("\n");
  const esc = host.replace(/[.]/g, "[.]");
  if (lines.some((l) => new RegExp(`^\\s*-?\\s*hostname:\\s*["']?${esc}["']?\\s*$`).test(l))) return { result: "present", text };
  const ingress = lines.findIndex((l) => /^ingress:\s*$/.test(l));
  if (ingress < 0) throw new Error("no ingress: section");
  // The catch-all is the last rule: "- service: ..." with no hostname.
  let catchAll = -1;
  for (let i = ingress + 1; i < lines.length; i++) {
    if (/^\S/.test(lines[i]) && lines[i].trim()) break;
    if (/^\s*-\s*service:/.test(lines[i])) catchAll = i;
  }
  if (catchAll < 0) throw new Error("no catch-all rule (- service: http_status:404) found");
  const indent = lines[catchAll].match(/^\s*/)[0];
  lines.splice(catchAll, 0, `${indent}- hostname: ${host}`, `${indent}  service: http://localhost:${port}`);
  return { result: "added", text: lines.join("\n") };
}

if (require.main === module) {
  const [file, host, port] = process.argv.slice(2);
  try {
    const out = addRoute(fs.readFileSync(file, "utf8"), host, port);
    if (out.result === "added") fs.writeFileSync(file, out.text);
    console.log(out.result);
  } catch (err) {
    console.error(err.message);
    process.exit(1);
  }
}
module.exports = { addRoute };

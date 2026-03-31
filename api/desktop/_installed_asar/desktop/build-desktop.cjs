const { spawn } = require("node:child_process");

function fmtTag(d = new Date()) {
  const p = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}-${p(d.getHours())}${p(d.getMinutes())}`;
}

const tag = process.env.BUILD_TAG || fmtTag();
const args = ["electron-builder", "--win", "nsis", "--x64"];

console.log(`[desktop-build] BUILD_TAG=${tag}`);
const child = spawn("npx", args, {
  stdio: "inherit",
  shell: true,
  env: { ...process.env, BUILD_TAG: tag },
});

child.on("exit", (code) => {
  process.exit(code || 0);
});


const fs = require("fs");
const path = require("path");

const dir = path.join(__dirname, "assets", "undercarriage");
const map = {
  overview: "overview.svg",
  bushings_links: "bushings-links.svg",
  track_shoe: "track-shoe.svg",
  carrier_rollers: "carrier-rollers.svg",
  track_rollers: "track-rollers.svg",
  track_sag: "track-sag.svg",
};

const lines = ["(function () {", "  const svg = {"];
for (const [key, file] of Object.entries(map)) {
  const raw = fs.readFileSync(path.join(dir, file), "utf8").trim();
  lines.push(`    ${key}: ${JSON.stringify(raw)},`);
}
lines.push("  };");
lines.push("  window.UC_DIAGRAM_SVG = svg;");
lines.push("})();");
lines.push("");

fs.writeFileSync(path.join(__dirname, "undercarriage-diagrams.js"), lines.join("\n"));
console.log("Wrote undercarriage-diagrams.js");

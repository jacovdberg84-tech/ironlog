// Reads a split web page's code the way the browser does: the scripts from
// one folder listed in the page's HTML, in load order, joined into one script.
import fs from "node:fs";

const webDir = new URL("../../web/", import.meta.url);

export function pageScriptFiles(page = "index.html", dir = "app") {
  const html = fs.readFileSync(new URL(page, webDir), "utf8");
  const re = new RegExp(`<script src="\\./(${dir}/[^"?]+)(?:\\?[^"]*)?"></script>`, "g");
  return [...html.matchAll(re)].map((m) => m[1]);
}

export function readPageSource(page = "index.html", dir = "app") {
  return pageScriptFiles(page, dir).map((f) => fs.readFileSync(new URL(f, webDir), "utf8")).join("\n");
}

export const webAppFiles = () => pageScriptFiles("index.html", "app");
export const readWebAppSource = () => readPageSource("index.html", "app");
export const maintenanceFiles = () => pageScriptFiles("maintenance.html", "maintenance");
export const readMaintenanceSource = () => readPageSource("maintenance.html", "maintenance");

// Reads the main web app the way the browser does: the web/app/*.js files
// listed in web/index.html, in load order, joined into one script.
import fs from "node:fs";

const webDir = new URL("../../web/", import.meta.url);

export function webAppFiles() {
  const html = fs.readFileSync(new URL("index.html", webDir), "utf8");
  return [...html.matchAll(/<script src="\.\/(app\/[^"?]+)(?:\?[^"]*)?"><\/script>/g)].map((m) => m[1]);
}

export function readWebAppSource() {
  return webAppFiles().map((f) => fs.readFileSync(new URL(f, webDir), "utf8")).join("\n");
}

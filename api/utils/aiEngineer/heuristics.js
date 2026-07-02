export function chooseTargetFiles(title, text) {
  const hay = `${String(title || "")} ${String(text || "")}`.toLowerCase();
  const files = new Set();
  if (/maint|service|mechanic|workshop|ai engineer|ai-engineer/.test(hay)) {
    files.add("web/maintenance.html");
    files.add("web/maintenance.js");
    files.add("api/routes/maintenance.routes.js");
    files.add("api/routes/aiEngineer.routes.js");
  }
  if (/lube|oil|dashboard/.test(hay)) {
    files.add("web/app.js");
    files.add("api/routes/dashboard.routes.js");
    files.add("api/routes/reports.routes.js");
  }
  if (/work.?order|breakdown/.test(hay)) {
    files.add("api/routes/workorders.routes.js");
    files.add("api/routes/breakdowns.routes.js");
  }
  if (/api|endpoint|route/.test(hay)) {
    files.add("api/server.js");
  }
  if (!files.size) {
    files.add("web/app.js");
    files.add("api/routes/maintenance.routes.js");
  }
  return [...files];
}

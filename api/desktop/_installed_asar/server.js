// IRONLOG/api/server.js
import Fastify from "fastify";
import cors from "@fastify/cors";
import breakdownRoutes from "./routes/breakdowns.routes.js";
import workOrderRoutes from "./routes/workorders.routes.js";
import stockRoutes from "./routes/stock.routes.js";
import assetRoutes from "./routes/assets.routes.js";
import hoursRoutes from "./routes/hours.routes.js";
import uploadRoutes from "./routes/upload.routes.js";
import kpiRoutes from "./routes/kpi.routes.js";
import maintenanceRoutes from "./routes/maintenance.routes.js";
import reportsRoutes from "./routes/reports.routes.js";
import alertsRoutes from "./routes/alerts.routes.js";
import dashboardRoutes from "./routes/dashboard.routes.js";
import authRoutes from "./routes/auth.routes.js";
import auditRoutes from "./routes/audit.routes.js";
import legalRoutes from "./routes/legal.routes.js";
import approvalsRoutes from "./routes/approvals.routes.js";
import procurementRoutes from "./routes/procurement.routes.js";
import operationsRoutes from "./routes/operations.routes.js";
import dispatchRoutes from "./routes/dispatch.routes.js";
import qualityRoutes from "./routes/quality.routes.js";
import syncRoutes from "./routes/sync.routes.js";

export function buildServer() {
  const isDesktopRuntime = process.env.IRONLOG_DESKTOP === "1";
  const app = Fastify({
    // In packaged Electron (Windows GUI), stdout/stderr can be unavailable.
    // Disable Fastify/Pino console logger there to avoid EBADF write crashes.
    logger: isDesktopRuntime ? false : true
  });

  app.register(cors, { origin: true });

  // Health check
  app.get("/health", async () => {
    return { ok: true, name: "IRONLOG", db: "sqlite" };
  });

  // Routes
  app.register(assetRoutes, { prefix: "/api/assets" });
  app.register(hoursRoutes, { prefix: "/api/hours" });
  app.register(uploadRoutes, { prefix: "/api/upload" });
  app.register(breakdownRoutes, { prefix: "/api/breakdowns" });
app.register(workOrderRoutes, { prefix: "/api/workorders" });
app.register(stockRoutes, { prefix: "/api/stock" });
app.register(kpiRoutes, { prefix: "/api/kpi" });
app.register(maintenanceRoutes, { prefix: "/api/maintenance" });
app.register(reportsRoutes, { prefix: "/api/reports" });
app.register(alertsRoutes, { prefix: "/api/alerts" });
app.register(dashboardRoutes, { prefix: "/api/dashboard" });
app.register(authRoutes, { prefix: "/api/auth" });
app.register(auditRoutes, { prefix: "/api/audit" });
app.register(legalRoutes, { prefix: "/api/legal" });
app.register(approvalsRoutes, { prefix: "/api/approvals" });
app.register(procurementRoutes, { prefix: "/api/procurement" });
app.register(operationsRoutes, { prefix: "/api/operations" });
app.register(dispatchRoutes, { prefix: "/api/dispatch" });
app.register(qualityRoutes, { prefix: "/api/quality" });
app.register(syncRoutes, { prefix: "/api/sync" });
  return app;
}
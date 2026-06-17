(function () {
  const A = window.IronlogAuth;
  if (!A) return;

  const ALLOWED_ROLES = ["artisan", "admin", "supervisor"];

  function qs(id) {
    return document.getElementById(id);
  }

  function showLogin(on) {
    qs("loginScreen")?.classList.toggle("hidden", !on);
    qs("appScreen")?.classList.toggle("hidden", on);
  }

  function statusLabel(s) {
    const v = String(s || "").toLowerCase();
    if (v === "assigned") return "Assigned";
    if (v === "in_progress") return "In progress";
    if (v === "completed") return "Awaiting approval";
    return v.replace(/_/g, " ");
  }

  function statusClass(s) {
    const v = String(s || "").toLowerCase();
    if (v === "assigned") return "assigned";
    if (v === "in_progress") return "in_progress";
    if (v === "completed") return "completed";
    return "";
  }

  function updateSessionChrome(user) {
    const label = qs("sessionLabel");
    if (!label) return;
    const name = user?.full_name || user?.username || A.getSessionUser();
    const role = A.getSessionRole();
    label.textContent = `${name} · ${role}`;
  }

  function ensureRoleAccess(user) {
    const roles = Array.isArray(user?.roles) ? user.roles : [user?.role || A.getSessionRole()];
    const ok = roles.some((r) => ALLOWED_ROLES.includes(String(r).toLowerCase()));
    if (!ok) {
      throw new Error("This terminal is for workshop technicians (artisan role). Ask your supervisor for access.");
    }
  }

  async function loadJobs() {
    const host = qs("jobsList");
    const summary = qs("jobsSummary");
    if (!host) return;
    host.innerHTML = `<div class="muted small">Loading…</div>`;
    const rows = await A.fetchJson(`${A.API}/api/workorders`);
    const list = Array.isArray(rows) ? rows : [];
    const active = list.filter((wo) => {
      const s = String(wo.status || "").toLowerCase();
      return ["assigned", "in_progress", "completed"].includes(s);
    });

    const counts = { assigned: 0, in_progress: 0, completed: 0 };
    active.forEach((wo) => {
      const s = String(wo.status || "").toLowerCase();
      if (counts[s] != null) counts[s] += 1;
    });

    if (summary) {
      summary.innerHTML = `
        <span class="pill assigned">Assigned: ${counts.assigned}</span>
        <span class="pill progress">In progress: ${counts.in_progress}</span>
        <span class="pill done">Submitted: ${counts.completed}</span>
      `;
    }

    if (!active.length) {
      host.innerHTML = `<div class="muted small">No jobs assigned to you right now.</div>`;
      return;
    }

    host.innerHTML = "";
    active.forEach((wo) => {
      const card = document.createElement("button");
      card.type = "button";
      card.className = "job-card";
      const st = String(wo.status || "").toLowerCase();
      card.innerHTML = `
        <div class="job-card-top">
          <span class="job-id">WO #${A.escapeHtml(wo.id)}</span>
          <span class="job-status ${statusClass(st)}">${A.escapeHtml(statusLabel(st))}</span>
        </div>
        <div class="job-asset">${A.escapeHtml(wo.asset_code || "-")} · ${A.escapeHtml(wo.asset_name || "")}</div>
        <div class="job-meta">${A.escapeHtml(wo.source || "")}${wo.opened_at ? ` · ${A.escapeHtml(wo.opened_at)}` : ""}</div>
      `;
      card.addEventListener("click", () => {
        window.location.href = `./workorder-qr.html?wo_id=${encodeURIComponent(wo.id)}`;
      });
      host.appendChild(card);
    });
  }

  function openWorkOrder() {
    const id = Number(qs("woLookupId")?.value || 0);
    if (!Number.isFinite(id) || id <= 0) {
      alert("Enter a valid work order number.");
      return;
    }
    window.location.href = `./workorder-qr.html?wo_id=${encodeURIComponent(id)}`;
  }

  async function handleLogin() {
    const errEl = qs("loginError");
    if (errEl) errEl.textContent = "";
    const username = String(qs("loginUsername")?.value || "").trim();
    const password = String(qs("loginPassword")?.value || "");
    const remember = qs("loginRemember")?.checked !== false;
    if (!username || !password) {
      if (errEl) errEl.textContent = "Enter username and password.";
      return;
    }
    try {
      const data = await A.login(username, password, remember);
      ensureRoleAccess(data.user);
      showLogin(false);
      updateSessionChrome(data.user);
      await loadJobs();
    } catch (e) {
      if (errEl) errEl.textContent = e.message || String(e);
    }
  }

  async function boot() {
    const user = await A.trySession();
    if (user) {
      try {
        ensureRoleAccess(user);
        showLogin(false);
        updateSessionChrome(user);
        await loadJobs();
        return;
      } catch (e) {
        A.clearSession();
        const errEl = qs("loginError");
        if (errEl) errEl.textContent = e.message || String(e);
      }
    }
    showLogin(true);
  }

  qs("loginBtn")?.addEventListener("click", () => handleLogin());
  qs("loginPassword")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") handleLogin();
  });
  qs("logoutBtn")?.addEventListener("click", () => {
    A.clearSession();
    showLogin(true);
    if (qs("loginPassword")) qs("loginPassword").value = "";
  });
  qs("refreshJobsBtn")?.addEventListener("click", () => loadJobs().catch((e) => alert(e.message || e)));
  qs("openWoBtn")?.addEventListener("click", openWorkOrder);
  qs("woLookupId")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") openWorkOrder();
  });

  boot().catch(() => showLogin(true));
})();

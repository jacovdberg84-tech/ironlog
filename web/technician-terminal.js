(function () {
  const A = window.IronlogAuth;
  if (!A) return;

  const ALLOWED_ROLES = ["artisan", "admin", "supervisor"];
  const MAX_PIN = 6;

  let selectedUsername = "";
  let pinValue = "";
  let roster = [];

  function qs(id) {
    return document.getElementById(id);
  }

  function showLogin(on) {
    qs("loginScreen")?.classList.toggle("hidden", !on);
    qs("appScreen")?.classList.toggle("hidden", on);
  }

  function showPinPanel(on) {
    qs("pinLoginPanel")?.classList.toggle("hidden", !on);
    qs("passwordLoginPanel")?.classList.toggle("hidden", on);
    qs("showPasswordLoginBtn")?.classList.toggle("hidden", !on);
  }

  function initials(label) {
    const parts = String(label || "").trim().split(/\s+/).filter(Boolean);
    if (!parts.length) return "?";
    if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
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

  function renderPinDisplay() {
    const el = qs("pinDisplay");
    if (!el) return;
    el.textContent = pinValue ? "•".repeat(pinValue.length) : "····";
  }

  function updatePinSubmitState() {
    const btn = qs("pinSubmitBtn");
    if (!btn) return;
    const ok = Boolean(selectedUsername) && pinValue.length >= 4;
    btn.disabled = !ok;
  }

  function selectUser(username, label) {
    selectedUsername = String(username || "").trim();
    pinValue = "";
    renderPinDisplay();
    document.querySelectorAll(".pin-user-tile").forEach((tile) => {
      tile.classList.toggle("selected", tile.dataset.username === selectedUsername);
    });
    const sel = qs("pinSelectedUser");
    if (sel) {
      sel.textContent = selectedUsername ? `Signing in as ${label || selectedUsername}` : "Select your name above";
      sel.classList.toggle("hidden", !selectedUsername);
    }
    updatePinSubmitState();
  }

  function renderRoster(list) {
    const host = qs("pinRoster");
    if (!host) return;
    roster = Array.isArray(list) ? list : [];
    if (!roster.length) {
      host.innerHTML = `<div class="muted small">No PIN users yet. Ask your supervisor to set a PIN in User Admin, or use username &amp; password below.</div>`;
      return;
    }
    host.innerHTML = "";
    roster.forEach((t) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "pin-user-tile";
      btn.dataset.username = t.username;
      btn.innerHTML = `
        <span class="pin-user-initials">${A.escapeHtml(initials(t.label))}</span>
        <span class="pin-user-label">${A.escapeHtml(t.label)}</span>
      `;
      btn.addEventListener("click", () => selectUser(t.username, t.label));
      host.appendChild(btn);
    });
  }

  async function loadPinRoster() {
    try {
      const list = await A.fetchPinRoster();
      renderRoster(list);
    } catch {
      renderRoster([]);
    }
  }

  function appendPinDigit(d) {
    if (!selectedUsername) {
      const err = qs("pinLoginError");
      if (err) err.textContent = "Select your name first.";
      return;
    }
    if (pinValue.length >= MAX_PIN) return;
    pinValue += String(d);
    renderPinDisplay();
    updatePinSubmitState();
    const err = qs("pinLoginError");
    if (err) err.textContent = "";
    if (pinValue.length >= 6) handlePinLogin();
  }

  function clearPin() {
    pinValue = "";
    renderPinDisplay();
    updatePinSubmitState();
  }

  function backPin() {
    pinValue = pinValue.slice(0, -1);
    renderPinDisplay();
    updatePinSubmitState();
  }

  async function completeLogin(user, remember) {
    ensureRoleAccess(user);
    showLogin(false);
    updateSessionChrome(user);
    await loadJobs();
  }

  async function handlePinLogin() {
    const errEl = qs("pinLoginError");
    if (errEl) errEl.textContent = "";
    if (!selectedUsername || pinValue.length < 4) return;
    const btn = qs("pinSubmitBtn");
    if (btn) btn.disabled = true;
    try {
      const data = await A.loginPin(selectedUsername, pinValue, false);
      clearPin();
      await completeLogin(data.user, false);
    } catch (e) {
      clearPin();
      if (errEl) errEl.textContent = e.message || String(e);
    } finally {
      updatePinSubmitState();
    }
  }

  async function handlePasswordLogin() {
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
      await completeLogin(data.user, remember);
    } catch (e) {
      if (errEl) errEl.textContent = e.message || String(e);
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

  async function boot() {
    if (!A.getAuthToken()) {
      showLogin(true);
      showPinPanel(true);
      await loadPinRoster();
      renderPinDisplay();
      return;
    }

    const user = await A.trySession();
    if (user) {
      try {
        await completeLogin(user, true);
        return;
      } catch (e) {
        A.clearSession();
        const errEl = qs("pinLoginError");
        if (errEl) errEl.textContent = e.message || String(e);
      }
    }
    showLogin(true);
    showPinPanel(true);
    await loadPinRoster();
    renderPinDisplay();
  }

  document.querySelectorAll(".pin-key[data-digit]").forEach((btn) => {
    btn.addEventListener("click", () => appendPinDigit(btn.dataset.digit));
  });
  qs("pinClearBtn")?.addEventListener("click", clearPin);
  qs("pinBackBtn")?.addEventListener("click", backPin);
  qs("pinSubmitBtn")?.addEventListener("click", () => handlePinLogin());
  qs("showPasswordLoginBtn")?.addEventListener("click", () => showPinPanel(false));
  qs("showPinLoginBtn")?.addEventListener("click", () => showPinPanel(true));
  qs("loginBtn")?.addEventListener("click", () => handlePasswordLogin());
  qs("loginPassword")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") handlePasswordLogin();
  });
  qs("logoutBtn")?.addEventListener("click", () => {
    A.clearSession();
    selectedUsername = "";
    clearPin();
    showLogin(true);
    showPinPanel(true);
    loadPinRoster();
    if (qs("loginPassword")) qs("loginPassword").value = "";
  });
  qs("refreshJobsBtn")?.addEventListener("click", () => loadJobs().catch((e) => alert(e.message || e)));
  qs("openWoBtn")?.addEventListener("click", openWorkOrder);
  qs("woLookupId")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") openWorkOrder();
  });

  boot().catch(() => {
    showLogin(true);
    showPinPanel(true);
  });
})();

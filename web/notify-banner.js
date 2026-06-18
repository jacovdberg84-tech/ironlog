/**
 * IRONLOG Notify — download banner for technician pages.
 */
(function () {
  const API = window.location?.origin || "http://localhost:3001";

  async function loadNotifyBanner() {
    const el = document.getElementById("notifyAppBanner");
    if (!el) return;
    try {
      const res = await fetch(`${API}/api/notifications/config`);
      if (!res.ok) return;
      const cfg = await res.json();
      const link = document.getElementById("notifyApkLink");
      if (link && cfg.apk_url) link.href = cfg.apk_url;
      const status = document.getElementById("notifyPushStatus");
      if (status) {
        status.textContent = cfg.push_enabled
          ? "Server push is configured."
          : "Install the app now; server push activates once Firebase is configured.";
      }
      el.style.display = "block";
    } catch {
      el.style.display = "block";
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", loadNotifyBanner);
  } else {
    loadNotifyBanner();
  }
})();

/**
 * IRONLOG Notify — download banner + Expo install QR for technician pages.
 */
(function () {
  const API = window.location?.origin || "http://localhost:3001";

  function qrImageUrl(targetUrl, size) {
    const s = size || 320;
    return `https://api.qrserver.com/v1/create-qr-code/?size=${s}x${s}&data=${encodeURIComponent(targetUrl)}`;
  }

  async function loadNotifyBanner() {
    const el = document.getElementById("notifyAppBanner");
    if (!el) return;
    try {
      const res = await fetch(`${API}/api/notifications/config`);
      if (!res.ok) return;
      const cfg = await res.json();

      const link = document.getElementById("notifyApkLink");
      if (link && cfg.apk_url) link.href = cfg.apk_url;

      const expoUrl = String(cfg.expo_install_url || "").trim();
      const expoLink = document.getElementById("notifyExpoLink");
      const qrImg = document.getElementById("notifyExpoQr");
      if (expoUrl) {
        if (expoLink) {
          expoLink.href = expoUrl;
          expoLink.textContent = "Open Expo install page";
        }
        if (qrImg) {
          qrImg.src = cfg.expo_qr_url || qrImageUrl(expoUrl, 320);
          qrImg.alt = "Scan to install IRONLOG Notify";
        }
      }

      const status = document.getElementById("notifyPushStatus");
      if (status) {
        status.textContent = cfg.push_enabled
          ? "Server push is configured. Scan the QR with your phone camera to install."
          : "Scan the QR to install; server push activates once Firebase is configured.";
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

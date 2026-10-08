// IRONLOG/web/app/wo-popup.js — open a work order over the current page.
// Part of the main app; index.html loads these files in order and they share one global scope.
//
// The pop-up shows the Work Orders page's own detail panel (workorders.html
// ?wo=ID&popup=1) in a frame, so everything on a job works the same as on the
// board: progress, parts, assign, complete, PDF. June opens it when you ask
// about a work order; anything else can call openWorkOrderPopup(id).

let woPopupEl = null;

function closeWorkOrderPopup() {
  if (!woPopupEl) return;
  woPopupEl.hidden = true;
  woPopupEl.querySelector("iframe")?.setAttribute("src", "about:blank");
  document.body.classList.remove("wo-popup-open");
}

function openWorkOrderPopup(id) {
  const woId = Math.trunc(Number(id));
  if (!Number.isFinite(woId) || woId <= 0) return;
  if (!woPopupEl) {
    woPopupEl = document.createElement("div");
    woPopupEl.className = "wo-popup-overlay";
    woPopupEl.hidden = true;
    woPopupEl.innerHTML = `
      <div class="wo-popup-dialog" role="dialog" aria-modal="true" aria-label="Work order">
        <div class="wo-popup-bar">
          <strong class="wo-popup-title"></strong>
          <span class="wo-popup-actions">
            <a class="wo-popup-full" target="_blank" rel="noopener">Open on Work Orders page</a>
            <button type="button" class="wo-popup-x" aria-label="Close work order">✕</button>
          </span>
        </div>
        <iframe class="wo-popup-frame" title="Work order detail"></iframe>
      </div>`;
    document.body.appendChild(woPopupEl);
    woPopupEl.addEventListener("click", (e) => {
      if (e.target === woPopupEl || e.target.closest(".wo-popup-x")) closeWorkOrderPopup();
    });
    document.addEventListener("keydown", (e) => { if (e.key === "Escape" && woPopupEl && !woPopupEl.hidden) closeWorkOrderPopup(); });
    window.addEventListener("message", (e) => {
      if (e.origin === window.location.origin && e.data?.type === "ironlog:wo-popup-close") closeWorkOrderPopup();
    });
  }
  woPopupEl.querySelector(".wo-popup-title").textContent = `WO #${woId}`;
  woPopupEl.querySelector(".wo-popup-full").href = `/web/workorders.html?wo=${woId}`;
  woPopupEl.querySelector("iframe").src = `/web/workorders.html?wo=${woId}&popup=1`;
  woPopupEl.hidden = false;
  document.body.classList.add("wo-popup-open");
}
window.openWorkOrderPopup = openWorkOrderPopup;

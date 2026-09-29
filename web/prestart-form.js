// IRONLOG/web/prestart-form.js — OK / Fault checklist shared by the machine
// and LDV pre-start pages. Every check needs an answer; a fault can carry a
// short comment that goes to the workshop with the repair work order.
(function () {
  const OPERATOR_KEY = "ironlog_prestart_operator";

  function esc(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // Same wording as the server's fault list ("Engine oil level", not "Engine oil level OK").
  function plainLabel(label) {
    return String(label || "").replace(/\s+OK$/i, "").trim() || String(label || "");
  }

  function itemHtml(it) {
    const key = esc(it.key);
    const label = plainLabel(it.label);
    return `<div class="pc-item" data-key="${key}" data-label="${esc(label)}">
      <div class="pc-label">${esc(label)}</div>
      <div class="pc-choice" role="group" aria-label="${esc(label)}">
        <button type="button" class="pc-ok" data-answer="ok" aria-pressed="false">OK</button>
        <button type="button" class="pc-fault" data-answer="fault" aria-pressed="false">Fault</button>
      </div>
      <input class="pc-comment" type="text" maxlength="300" placeholder="What is wrong? (optional)" hidden />
    </div>`;
  }

  function setState(item, answer) {
    item.dataset.state = answer || "";
    item.querySelectorAll("[data-answer]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.answer === answer)));
    const comment = item.querySelector(".pc-comment");
    if (comment) comment.hidden = answer !== "fault";
  }

  /** sections: [{ title, items: [{ key, label }] }] */
  function render(root, sections, onChange) {
    if (!root) return;
    root.innerHTML = sections.map((sec) => `
      <section class="pc-section">
        ${sec.title ? `<h3 class="pc-section-title">${esc(sec.title)}</h3>` : ""}
        ${(sec.items || []).filter((it) => String(it.key || "").trim()).map(itemHtml).join("")}
      </section>`).join("");
    if (root.dataset.bound) return;
    root.dataset.bound = "1";
    root.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-answer]");
      if (!btn) return;
      const item = btn.closest(".pc-item");
      item.classList.remove("is-missing");
      setState(item, btn.dataset.answer);
      if (btn.dataset.answer === "fault") item.querySelector(".pc-comment")?.focus();
      root.dispatchEvent(new CustomEvent("pc-change"));
    });
    root.addEventListener("pc-change", () => onChange && onChange(progress(root)));
  }

  function items(root) {
    return root ? [...root.querySelectorAll(".pc-item")] : [];
  }

  function progress(root) {
    const all = items(root);
    return {
      total: all.length,
      answered: all.filter((i) => i.dataset.state).length,
      faults: all.filter((i) => i.dataset.state === "fault").length,
    };
  }

  /** { checklist: {key: bool}, faults: {key: comment}, unanswered: [label] } */
  function read(root) {
    const checklist = {};
    const faults = {};
    const unanswered = [];
    for (const item of items(root)) {
      const key = item.dataset.key;
      const st = item.dataset.state;
      if (!st) {
        unanswered.push(item.dataset.label || key);
        continue;
      }
      checklist[key] = st === "ok";
      if (st === "fault") faults[key] = String(item.querySelector(".pc-comment")?.value || "").trim();
    }
    return { checklist, faults, unanswered };
  }

  /** Accepts the saved checklist as an array of {key, ok} or an object {key: bool}. */
  function apply(root, saved) {
    const byKey = {};
    if (Array.isArray(saved)) saved.forEach((r) => { byKey[String(r?.key || "")] = r?.ok; });
    else if (saved && typeof saved === "object") Object.assign(byKey, saved);
    for (const item of items(root)) {
      const v = byKey[item.dataset.key];
      setState(item, v === true ? "ok" : v === false ? "fault" : "");
    }
    root?.dispatchEvent(new CustomEvent("pc-change"));
  }

  function firstUnanswered(root) {
    return items(root).find((i) => !i.dataset.state) || null;
  }

  /** Fills the operator name from last time and remembers it on change. */
  function rememberOperator(input) {
    if (!input) return;
    try {
      if (!input.value) input.value = localStorage.getItem(OPERATOR_KEY) || "";
    } catch {}
    input.addEventListener("change", () => {
      try { localStorage.setItem(OPERATOR_KEY, String(input.value || "").trim()); } catch {}
    });
  }

  // Saved notes carry a "Faults: Label (comment); ..." part added by the server.
  // Split it off so resubmitting does not repeat it, and put comments back.
  function splitSavedNotes(notes) {
    const parts = String(notes || "").split(" | ").map((x) => x.trim()).filter(Boolean);
    const faultPart = parts.find((x) => x.startsWith("Faults: "));
    const comments = {};
    if (faultPart) {
      for (const bit of faultPart.slice(8).split("; ")) {
        const m = bit.match(/^(.+?) \((.*)\)$/);
        if (m) comments[m[1]] = m[2];
      }
    }
    return {
      notes: parts.filter((x) => !x.startsWith("Faults: ") && x !== "KM flagged for supervisor review").join(" | "),
      comments,
    };
  }

  function applyFaultComments(root, commentsByLabel) {
    for (const item of items(root)) {
      const c = commentsByLabel[item.dataset.label];
      const input = item.querySelector(".pc-comment");
      if (input) input.value = c && item.dataset.state === "fault" ? c : "";
    }
  }

  window.IronlogPrestart = { render, read, apply, progress, firstUnanswered, rememberOperator, splitSavedNotes, applyFaultComments };
})();

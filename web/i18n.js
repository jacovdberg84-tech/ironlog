// IRONLOG/web/i18n.js — English / Portuguese on every desk page.
//
// The pages are written in English. When Portuguese is chosen, this file swaps every
// known English phrase on the page for its Portuguese text (web/i18n-pt.js): text,
// placeholders, tooltips and button labels, including what the screens draw later
// (a watcher translates new content as it appears). Unknown text, such as machine
// codes, names and descriptions people typed, is left exactly as it is. Switching back
// to English restores the original text. Phrases with numbers use {n}-style patterns.
//
// The choice is kept in localStorage ("ironlog_lang"), shared by every desk page; the
// technician portal keeps its own switch. Mark an element data-no-translate to leave it alone.
(function (global) {
  const KEY = "ironlog_lang";
  const SKIP = new Set(["SCRIPT", "STYLE", "CODE", "PRE", "TEXTAREA", "NOSCRIPT", "SVG"]);
  const ATTRS = ["placeholder", "title", "aria-label"];
  const NUMERIC = new Set(["n", "h", "a", "b"]);
  const dicts = {};
  const patterns = {};
  const original = new WeakMap(); // text node → its English text
  const written = new WeakSet(); // text nodes this file just changed (the watcher ignores them)

  const norm = (s) => String(s).replace(/\s+/g, " ").trim();
  const escapeRe = (s) => s.replace(/[.*+?^$()|[\]\\]/g, "\\$&");

  /** Adds phrases for a language. A key with {name} parts becomes a pattern. */
  function add(lang, entries) {
    dicts[lang] = dicts[lang] || new Map();
    patterns[lang] = patterns[lang] || [];
    for (const [en, tr] of Object.entries(entries)) {
      if (/\{\w+\}/.test(en)) {
        const names = [...en.matchAll(/\{(\w+)\}/g)].map((m) => m[1]);
        // {n}, {h}, {a} and {b} are numbers; other names take any words.
        const parts = en.split(/\{\w+\}/).map(escapeRe);
        const re = new RegExp(`^${parts.map((p, i) => (i ? (NUMERIC.has(names[i - 1]) ? "([\\d.,]+)" : "(.+?)") : "") + p).join("")}$`);
        patterns[lang].push({ re, names, tr, fixed: en.replace(/\{\w+\}/g, "").length });
      } else {
        dicts[lang].set(norm(en), tr);
      }
    }
    // The pattern with the most fixed words wins ("Due @ {h}h" before "Due {d}").
    patterns[lang].sort((a, b) => b.fixed - a.fixed);
  }

  function lang() {
    try {
      // Someone who chose Portuguese in the technician portal or a pre-start form starts in Portuguese here too.
      const v = localStorage.getItem(KEY) || (localStorage.getItem("ironlog_prestart_lang") === "pt" ? "pt" : "en");
      return String(v).toLowerCase();
    } catch {
      return "en";
    }
  }

  /** The phrase in the chosen language (unknown phrases come back unchanged). */
  function t(text, l = lang()) {
    const src = norm(text);
    if (!src || l === "en" || !dicts[l]) return text;
    const hit = dicts[l].get(src);
    if (hit != null) return hit;
    for (const p of patterns[l] || []) {
      const m = src.match(p.re);
      if (!m) continue;
      return p.tr.replace(/\{(\w+)\}/g, (_, name) => {
        const i = p.names.indexOf(name);
        return i >= 0 ? t(m[i + 1], l) : `{${name}}`;
      });
    }
    return text;
  }

  function skipped(el, { attrs = false } = {}) {
    for (let e = el; e && e.nodeType === 1; e = e.parentElement) {
      // A text box's typed text is left alone, but its placeholder and tooltip are translated.
      const tag = e.tagName.toUpperCase();
      if (attrs && e === el && tag === "TEXTAREA") continue;
      if (SKIP.has(tag) || e.isContentEditable || e.hasAttribute("data-no-translate")) return true;
    }
    return false;
  }

  function translateText(node, l) {
    if (!original.has(node)) original.set(node, node.data);
    const en = original.get(node);
    let next = en;
    if (l !== "en") {
      const core = norm(en);
      if (!core) return;
      const tr = t(core, l);
      // Unknown phrases stay (or go back to) the English text.
      if (tr !== core) next = `${en.match(/^\s*/)[0]}${tr}${en.match(/\s*$/)[0]}`;
    }
    if (node.data !== next) {
      written.add(node);
      node.data = next;
    }
  }

  function translateAttrs(el, l) {
    for (const a of ATTRS) {
      if (!el.hasAttribute(a)) continue;
      const store = `data-i18n-${a}`;
      if (!el.hasAttribute(store)) el.setAttribute(store, el.getAttribute(a));
      const en = el.getAttribute(store);
      const next = l === "en" ? en : t(en, l);
      if (el.getAttribute(a) !== next) el.setAttribute(a, next);
    }
    // Buttons and inputs whose label is their value.
    if (el.tagName === "INPUT" && ["button", "submit"].includes(el.type)) {
      if (!el.hasAttribute("data-i18n-value")) el.setAttribute("data-i18n-value", el.value);
      el.value = l === "en" ? el.getAttribute("data-i18n-value") : t(el.getAttribute("data-i18n-value"), l);
    }
  }

  /** Translates everything under root into the chosen language (or back to English). */
  function apply(root = document.body, l = lang()) {
    if (!root) return;
    if (root.nodeType === 3) {
      if (!skipped(root.parentElement)) translateText(root, l);
      return;
    }
    if (root.nodeType !== 1) return;
    if (!skipped(root, { attrs: true })) translateAttrs(root, l);
    if (skipped(root)) return;
    root.querySelectorAll("[placeholder],[title],[aria-label],input[type=button],input[type=submit]").forEach((el) => { if (!skipped(el, { attrs: true })) translateAttrs(el, l); });
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: (n) => (n.data.trim() && !skipped(n.parentElement) ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT),
    });
    const nodes = [];
    while (walker.nextNode()) nodes.push(walker.currentNode);
    nodes.forEach((n) => translateText(n, l));
    document.documentElement.lang = l === "pt" ? "pt" : "en";
  }

  // New content (lists, cards, messages drawn by the screens) is translated as it appears.
  let pending = new Set();
  let scheduled = false;
  function flush() {
    scheduled = false;
    const l = lang();
    const batch = pending;
    pending = new Set();
    if (l === "en") return;
    batch.forEach((n) => { if (n.isConnected) apply(n, l); });
  }
  function watch() {
    if (!global.MutationObserver || !document.body) return;
    new MutationObserver((records) => {
      if (lang() === "en") return;
      for (const r of records) {
        if (r.type === "characterData") {
          if (written.has(r.target)) { written.delete(r.target); continue; }
          original.delete(r.target); // the screen wrote new English text
          pending.add(r.target);
        } else if (r.type === "attributes") {
          const a = r.attributeName;
          if (ATTRS.includes(a)) {
            const el = r.target;
            const now = el.getAttribute(a);
            if (now !== t(el.getAttribute(`data-i18n-${a}`) || "")) {
              el.setAttribute(`data-i18n-${a}`, now);
              pending.add(el);
            }
          }
        } else {
          r.addedNodes.forEach((n) => pending.add(n));
        }
      }
      if (pending.size && !scheduled) {
        scheduled = true;
        (global.requestAnimationFrame || setTimeout)(flush);
      }
    }).observe(document.body, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ATTRS });
  }

  function setLang(l) {
    try { localStorage.setItem(KEY, l); } catch { /* private mode: this page only */ }
    apply(document.body, l);
    document.querySelectorAll("[data-i18n-switch]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.i18nSwitch === l)));
    global.dispatchEvent?.(new CustomEvent("ironlog:lang", { detail: { lang: l } }));
  }

  /** A small EN | PT switch (used in the shared menu). */
  function switchHtml() {
    const l = lang();
    return `<div class="lang-switch" role="group" aria-label="Language / Idioma" data-no-translate>
      <button type="button" data-i18n-switch="en" aria-pressed="${l === "en"}">EN</button>
      <button type="button" data-i18n-switch="pt" aria-pressed="${l === "pt"}">PT</button>
    </div>`;
  }

  if (typeof document !== "undefined") {
    document.addEventListener("click", (e) => {
      const b = e.target?.closest?.("[data-i18n-switch]");
      if (b) setLang(b.dataset.i18nSwitch);
    });
    const start = () => { watch(); if (lang() !== "en") apply(); };
    if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start);
    else start();
  }

  /** Whether a phrase has a translation (some read the same in both languages, like "Stock:"). */
  function has(text, l) {
    const src = norm(text);
    return Boolean(dicts[l]?.has(src)) || (patterns[l] || []).some((p) => p.re.test(src));
  }

  global.IronlogI18n = { add, t, has, apply, lang, setLang, switchHtml };
})(typeof window !== "undefined" ? window : globalThis);

// IRONLOG/web/prestart-form.js — OK / Fault checklist shared by the machine
// and LDV pre-start pages. Every check needs an answer; a fault can carry a
// short comment that goes to the workshop with the repair work order.
(function () {
  const OPERATOR_KEY = "ironlog_prestart_operator";
  const LANG_KEY = "ironlog_prestart_lang";

  // Operator-facing wording. Checks come from the server with label / label_pt.
  const TEXT = {
    en: {
      yourName: "Your name", operatorPh: "Operator name", driverPh: "Driver name",
      hourMeter: "Hour meter (optional)", odometerOpt: "Odometer km (optional)", odometer: "Odometer (km) *",
      odometerPh: "From the dashboard", lastReading: "Last reading", date: "Date",
      tapEvery: "Tap OK or Fault for every check.", ok: "OK", fault: "Fault", whatWrong: "What is wrong? (optional)",
      progress: "{a} of {t} checked", faultCount: "{n} fault(s)",
      submit: "Submit pre-start", submitFaults: "Submit and report faults",
      addNote: "Add a note or photo (optional)", note: "Note", notePh: "Anything the workshop should know",
      photo: "Photo", photoHint: "The photo is attached when you submit.",
      moreChecks: "Tap OK or Fault for {n} more check(s).", enterName: "Enter your name.",
      enterKm: "Enter the odometer km.", badNumber: "Enter a number.",
      doneTitle: "Pre-start done", doneText: "{asset} is checked for {date}. Safe working!",
      faultsTitle: "Faults reported", faultsText: "{n} fault(s) sent to the workshop{wo}: {list}. Tell your foreman before you operate.",
      woRef: " (work order #{id})",
      savedPhone: "Saved on this phone", savedPhoneText: "No signal: the pre-start will send automatically when you are back online.",
      offlineFault: " You marked faults — tell your foreman before you operate.",
      already: "You already submitted this pre-start today. Change any answer and submit again if needed.",
      kmTitle: "Saved — check the km", kmText: "The km reading looks unusual. A supervisor will check it.",
      openPdf: "Open check PDF", changeAnswers: "Change answers", machineDetails: "Machine details",
      vehicleDetails: "Vehicle details", reload: "Reload", loading: "Loading…",
      photoAttached: " Photo attached.", photoNot: " Photo not attached: ",
      online: "Online.", offline: "No signal.", queued: "{n} pre-start(s) waiting to send.",
      vehiclePrestart: "Vehicle pre-start", machinePrestart: "Machine pre-start",
      noTemplate: "This machine has no pre-start checklist yet. Tell the workshop.",
      noticeTitle: "Message from the workshop", noticeAck: "I have read this message",
      noticeMust: "Tick 'I have read this message' at the top first.", noticeWo: "Work order #{id}",
    },
    pt: {
      yourName: "O seu nome", operatorPh: "Nome do operador", driverPh: "Nome do motorista",
      hourMeter: "Horímetro (opcional)", odometerOpt: "Conta-quilómetros km (opcional)", odometer: "Conta-quilómetros (km) *",
      odometerPh: "Do painel", lastReading: "Última leitura", date: "Data",
      tapEvery: "Toque em OK ou Avaria em cada verificação.", ok: "OK", fault: "Avaria", whatWrong: "O que está mal? (opcional)",
      progress: "{a} de {t} verificados", faultCount: "{n} avaria(s)",
      submit: "Enviar pré-arranque", submitFaults: "Enviar e comunicar avarias",
      addNote: "Adicionar nota ou foto (opcional)", note: "Nota", notePh: "Algo que a oficina deva saber",
      photo: "Foto", photoHint: "A foto é anexada quando enviar.",
      moreChecks: "Toque em OK ou Avaria em mais {n} verificação(ões).", enterName: "Escreva o seu nome.",
      enterKm: "Escreva os km do conta-quilómetros.", badNumber: "Escreva um número.",
      doneTitle: "Pré-arranque concluído", doneText: "{asset} verificado em {date}. Bom trabalho em segurança!",
      faultsTitle: "Avarias comunicadas", faultsText: "{n} avaria(s) enviada(s) à oficina{wo}: {list}. Informe o seu encarregado antes de operar.",
      woRef: " (ordem de trabalho n.º {id})",
      savedPhone: "Guardado neste telemóvel", savedPhoneText: "Sem rede: o pré-arranque será enviado automaticamente quando houver rede.",
      offlineFault: " Marcou avarias — informe o seu encarregado antes de operar.",
      already: "Já enviou este pré-arranque hoje. Altere as respostas e envie de novo se necessário.",
      kmTitle: "Guardado — verifique os km", kmText: "A leitura dos km parece invulgar. Um supervisor vai verificar.",
      openPdf: "Abrir PDF", changeAnswers: "Alterar respostas", machineDetails: "Detalhes da máquina",
      vehicleDetails: "Detalhes do veículo", reload: "Recarregar", loading: "A carregar…",
      photoAttached: " Foto anexada.", photoNot: " Foto não anexada: ",
      online: "Com rede.", offline: "Sem rede.", queued: "{n} pré-arranque(s) à espera de envio.",
      vehiclePrestart: "Pré-arranque do veículo", machinePrestart: "Pré-arranque da máquina",
      noTemplate: "Esta máquina ainda não tem lista de pré-arranque. Informe a oficina.",
      noticeTitle: "Mensagem da oficina", noticeAck: "Li esta mensagem",
      noticeMust: "Marque primeiro 'Li esta mensagem' no topo.", noticeWo: "Ordem de trabalho n.º {id}",
    },
  };

  function lang() {
    try {
      const saved = localStorage.getItem(LANG_KEY);
      if (saved === "en" || saved === "pt") return saved;
    } catch {}
    return String(navigator.language || "").toLowerCase().startsWith("pt") ? "pt" : "en";
  }

  function t(key, vars = {}) {
    const s = (TEXT[lang()] && TEXT[lang()][key]) || TEXT.en[key] || key;
    return s.replace(/\{(\w+)\}/g, (_, k) => (vars[k] == null ? "" : String(vars[k])));
  }

  /** Static page text: data-i18n (text), data-i18n-ph (placeholder). */
  function translatePage(doc = document) {
    doc.documentElement.lang = lang();
    doc.querySelectorAll("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
    doc.querySelectorAll("[data-i18n-ph]").forEach((el) => { el.placeholder = t(el.dataset.i18nPh); });
    doc.querySelectorAll("[data-lang]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.lang === lang())));
  }

  /** EN | PT buttons (elements with data-lang). onChange runs after the switch. */
  function bindLanguageToggle(onChange) {
    document.querySelectorAll("[data-lang]").forEach((b) => {
      b.addEventListener("click", () => {
        try { localStorage.setItem(LANG_KEY, b.dataset.lang); } catch {}
        translatePage();
        onChange && onChange(lang());
      });
    });
    translatePage();
  }

  function esc(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // Same wording as the server's fault list ("Engine oil level", not "Engine oil level OK").
  function plainLabel(label) {
    return String(label || "").replace(/\s+OK$/i, "").trim() || String(label || "");
  }

  // In Portuguese the check reads in Portuguese with the English underneath, so
  // a foreman or artisan reading over the operator's shoulder can follow it too.
  function labelHtml(en, pt) {
    if (lang() === "pt" && pt) return `${esc(plainLabel(pt))}<small class="pc-label-en">${esc(plainLabel(en))}</small>`;
    return esc(plainLabel(en));
  }

  function itemHtml(it) {
    const key = esc(it.key);
    const label = plainLabel(it.label);
    return `<div class="pc-item" data-key="${key}" data-label="${esc(label)}">
      <div class="pc-label">${labelHtml(it.label, it.label_pt)}</div>
      <div class="pc-choice" role="group" aria-label="${esc(label)}">
        <button type="button" class="pc-ok" data-answer="ok" aria-pressed="false">${esc(t("ok"))}</button>
        <button type="button" class="pc-fault" data-answer="fault" aria-pressed="false">${esc(t("fault"))}</button>
      </div>
      <input class="pc-comment" type="text" maxlength="300" placeholder="${esc(t("whatWrong"))}" hidden />
    </div>`;
  }

  function setState(item, answer) {
    item.dataset.state = answer || "";
    item.querySelectorAll("[data-answer]").forEach((b) => b.setAttribute("aria-pressed", String(b.dataset.answer === answer)));
    const comment = item.querySelector(".pc-comment");
    if (comment) comment.hidden = answer !== "fault";
  }

  /** sections: [{ title, title_pt?, items: [{ key, label, label_pt? }] }] */
  function render(root, sections, onChange) {
    if (!root) return;
    // Re-rendering (e.g. after a language switch) keeps the operator's answers.
    const before = root.querySelector(".pc-item") ? read(root) : null;
    root.innerHTML = sections.map((sec) => `
      <section class="pc-section">
        ${sec.title ? `<h3 class="pc-section-title">${esc(lang() === "pt" && sec.title_pt ? sec.title_pt : sec.title)}</h3>` : ""}
        ${(sec.items || []).filter((it) => String(it.key || "").trim()).map(itemHtml).join("")}
      </section>`).join("");
    root._pcOnChange = onChange;
    if (before) {
      for (const item of items(root)) {
        const k = item.dataset.key;
        if (k in before.checklist) setState(item, before.checklist[k] ? "ok" : "fault");
        const c = item.querySelector(".pc-comment");
        if (c && before.faults[k]) c.value = before.faults[k];
      }
      root.dispatchEvent(new CustomEvent("pc-change"));
    }
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
    root.addEventListener("pc-change", () => root._pcOnChange && root._pcOnChange(progress(root)));
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

  /**
   * Workshop notice for this machine (from the server context): shown in red
   * above the checklist in the operator's language, with an "I have read this"
   * tick that must be set before submitting.
   */
  function renderNotice(host, notice) {
    if (!host) return;
    host.dataset.noticeId = notice && notice.id ? String(notice.id) : "";
    host.__notice = notice || null;
    if (!notice || !notice.id) {
      host.innerHTML = "";
      host.hidden = true;
      return;
    }
    const pt = lang() === "pt";
    const main = (pt && notice.message_pt) || notice.message_en;
    const other = pt ? (notice.message_pt ? notice.message_en : "") : notice.message_pt || "";
    const wasTicked = host.querySelector(".pc-wn-ack input")?.checked;
    host.hidden = false;
    host.innerHTML = `
      <div class="pc-wn" role="alert">
        <div class="pc-wn-title">⚠ ${esc(t("noticeTitle"))}</div>
        <div class="pc-wn-text">${esc(main)}</div>
        ${other ? `<div class="pc-wn-other">${esc(other)}</div>` : ""}
        ${notice.work_order_id ? `<div class="pc-wn-wo">${esc(t("noticeWo", { id: notice.work_order_id }))}</div>` : ""}
        <label class="pc-wn-ack"><input type="checkbox" ${wasTicked ? "checked" : ""} /> ${esc(t("noticeAck"))}</label>
      </div>`;
  }

  /** The notice id to send when ticked; throws when a notice is shown but not ticked. */
  function noticeAck(host) {
    if (!host || !host.dataset.noticeId) return null;
    const ticked = host.querySelector(".pc-wn-ack input")?.checked;
    if (!ticked) {
      host.scrollIntoView({ behavior: "smooth", block: "start" });
      host.classList.add("is-missing");
      throw new Error(t("noticeMust"));
    }
    host.classList.remove("is-missing");
    return Number(host.dataset.noticeId);
  }

  window.IronlogPrestart = {
    render, read, apply, progress, firstUnanswered, rememberOperator, splitSavedNotes, applyFaultComments,
    lang, t, translatePage, bindLanguageToggle, renderNotice, noticeAck,
  };
})();

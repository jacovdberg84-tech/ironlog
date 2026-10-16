// IRONLOG/api/utils/failureCodes.js — why machines fail, counted the same way every time.
//
// Every breakdown gets four codes picked from lists, in the ISO 14224 style that
// reliability engineers expect: the component (system) that failed, the problem
// seen (failure mode), the cause, and the action taken. Free text stays for detail.
// The technician suggests component and problem when completing the job in the
// portal; the foreman checks them and adds cause and action. Codes are confirmed
// when all four are set. Lists are company-wide, in English and Portuguese, and an
// administrator or foreman can add, rename or switch off entries (a code in use is
// never deleted, so old breakdowns keep their meaning).

export const FAILURE_KINDS = {
  component: "Component",
  problem: "Problem",
  cause: "Cause",
  action: "Action",
};

// [code, English, Portuguese]
export const DEFAULT_FAILURE_CODES = {
  component: [
    ["ENG", "Engine", "Motor"],
    ["FUEL", "Fuel system", "Sistema de combustível"],
    ["COOL", "Cooling system", "Sistema de arrefecimento"],
    ["AIRX", "Air intake and exhaust", "Admissão de ar e escape"],
    ["ELEC", "Electrical", "Sistema eléctrico"],
    ["STRT", "Starting and charging", "Arranque e carga"],
    ["HYD", "Hydraulics", "Hidráulica"],
    ["TRANS", "Transmission", "Transmissão"],
    ["FDRV", "Final drive / differential", "Redução final / diferencial"],
    ["BRK", "Brakes", "Travões"],
    ["STR", "Steering", "Direcção"],
    ["SUSP", "Suspension", "Suspensão"],
    ["TYRE", "Tyres and wheels", "Pneus e jantes"],
    ["UC", "Undercarriage", "Material rodante"],
    ["GET", "G.E.T. and bucket", "G.E.T. e balde"],
    ["BOOM", "Boom, arm and cylinders", "Lança, braço e cilindros"],
    ["SWG", "Swing system", "Sistema de giro"],
    ["FRAME", "Frame and body", "Chassis e carroçaria"],
    ["CAB", "Cab and operator controls", "Cabine e comandos"],
    ["AC", "Air-conditioning", "Ar condicionado"],
    ["LUBE", "Auto-lube system", "Sistema de lubrificação automática"],
    ["SAFE", "Safety systems (fire, lights, alarms)", "Sistemas de segurança (incêndio, luzes, alarmes)"],
    ["OTHER", "Other", "Outro"],
  ],
  problem: [
    ["LEAK", "Leak (oil, fuel, coolant, air)", "Fuga (óleo, combustível, refrigerante, ar)"],
    ["BURST", "Burst hose or pipe", "Mangueira ou tubo rebentado"],
    ["OHT", "Overheating", "Sobreaquecimento"],
    ["NOST", "Will not start", "Não arranca"],
    ["LPWR", "Low power or slow", "Pouca potência ou lento"],
    ["NOISE", "Abnormal noise or vibration", "Ruído ou vibração anormal"],
    ["BRKN", "Broken or cracked", "Partido ou rachado"],
    ["WEAR", "Worn out", "Desgastado"],
    ["ELF", "Electrical fault / no power", "Avaria eléctrica / sem energia"],
    ["ALARM", "Warning light or fault code", "Luz de aviso ou código de avaria"],
    ["STUCK", "Stuck or will not move", "Preso ou não se move"],
    ["PUNC", "Puncture or tyre damage", "Furo ou pneu danificado"],
    ["CONT", "Contaminated (fluid or filter)", "Contaminado (fluido ou filtro)"],
    ["ADJ", "Out of adjustment", "Desregulado"],
    ["OTHER", "Other", "Outro"],
  ],
  cause: [
    ["WEAR", "Normal wear and tear", "Desgaste normal"],
    ["MAINT", "Missed or late maintenance", "Manutenção falhada ou atrasada"],
    ["LUBE", "Lack of lubrication", "Falta de lubrificação"],
    ["OPER", "Operator error or misuse", "Erro ou mau uso do operador"],
    ["ACC", "Accident or impact damage", "Acidente ou dano por impacto"],
    ["OVER", "Overload", "Sobrecarga"],
    ["CONT", "Contamination (dirt, water)", "Contaminação (sujidade, água)"],
    ["ENV", "Site conditions (heat, dust, mud)", "Condições do local (calor, pó, lama)"],
    ["PART", "Faulty part fitted", "Peça defeituosa montada"],
    ["INST", "Poor workmanship or fitting", "Má montagem ou mão-de-obra"],
    ["DESIGN", "Design or manufacturing fault", "Defeito de projecto ou fabrico"],
    ["UNK", "Not known", "Desconhecida"],
    ["OTHER", "Other", "Outra"],
  ],
  action: [
    ["REPL", "Replaced part", "Peça substituída"],
    ["REP", "Repaired", "Reparado"],
    ["ADJ", "Adjusted or reset", "Ajustado ou reposto"],
    ["CLEAN", "Cleaned or flushed", "Limpo ou lavado"],
    ["TIGHT", "Tightened or secured", "Apertado ou fixado"],
    ["REFILL", "Topped up or refilled", "Atestado ou reabastecido"],
    ["WELD", "Welded or fabricated", "Soldado ou fabricado"],
    ["OVH", "Overhauled", "Revisão geral"],
    ["SERV", "Serviced", "Manutenção feita"],
    ["NFF", "Inspected, no fault found", "Inspeccionado, sem avaria"],
    ["MOD", "Modified or upgraded", "Modificado ou melhorado"],
    ["OTHER", "Other", "Outro"],
  ],
};

const CODE = /^[A-Z0-9][A-Z0-9_-]{0,15}$/;
const COLS = ["fc_component", "fc_problem", "fc_cause", "fc_action", "fc_suggested_by", "fc_suggested_at", "fc_confirmed_by", "fc_confirmed_at"];

function hasColumn(db, table, col) {
  try {
    return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
  } catch {
    return false;
  }
}

/** The code lists (filled with the defaults the first time) and the breakdown columns. */
export function ensureFailureCodeTables(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS failure_codes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      kind TEXT NOT NULL CHECK (kind IN ('component', 'problem', 'cause', 'action')),
      code TEXT NOT NULL,
      label_en TEXT NOT NULL,
      label_pt TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      sort INTEGER NOT NULL DEFAULT 0,
      updated_at TEXT,
      updated_by TEXT,
      UNIQUE (kind, code)
    );
  `);
  if (!db.prepare(`SELECT 1 FROM failure_codes LIMIT 1`).get()) {
    const ins = db.prepare(`INSERT INTO failure_codes (kind, code, label_en, label_pt, sort) VALUES (?, ?, ?, ?, ?)`);
    db.transaction(() => {
      for (const [kind, rows] of Object.entries(DEFAULT_FAILURE_CODES)) rows.forEach(([code, en, pt], i) => ins.run(kind, code, en, pt, (i + 1) * 10));
    })();
  }
  if (db.prepare(`SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'breakdowns'`).get()) {
    for (const col of COLS) if (!hasColumn(db, "breakdowns", col)) db.prepare(`ALTER TABLE breakdowns ADD COLUMN ${col} TEXT`).run();
  }
}

/** { component: [{code, label_en, label_pt, active}], … } in list order; active only unless all. */
export function listFailureCodes(db, { all = false } = {}) {
  ensureFailureCodeTables(db);
  const out = Object.fromEntries(Object.keys(FAILURE_KINDS).map((k) => [k, []]));
  for (const r of db.prepare(`SELECT kind, code, label_en, label_pt, active FROM failure_codes ${all ? "" : "WHERE active = 1"} ORDER BY sort, label_en`).all()) {
    out[r.kind].push({ code: r.code, label_en: r.label_en, label_pt: r.label_pt || "", active: Boolean(r.active) });
  }
  return out;
}

/**
 * Saves one list as edited: [{ code, label_en, label_pt, active }] in order. Codes are kept (never
 * deleted); a code missing from the list is switched off. Returns { ok, list } or { ok:false, errors }.
 */
export function saveFailureCodeList(db, kind, rows, by = "") {
  ensureFailureCodeTables(db);
  if (!FAILURE_KINDS[kind]) return { ok: false, errors: ["Unknown list."] };
  if (!Array.isArray(rows) || !rows.length) return { ok: false, errors: ["A list needs at least one entry."] };
  if (rows.length > 200) return { ok: false, errors: ["At most 200 entries in a list."] };
  const errors = [];
  const seen = new Set();
  const clean = rows.map((r, i) => {
    const code = String(r?.code || "").trim().toUpperCase();
    const en = String(r?.label_en || "").trim();
    const pt = String(r?.label_pt || "").trim();
    if (!CODE.test(code)) errors.push(`Line ${i + 1}: the code must be 1-16 letters or numbers, like HYD.`);
    else if (seen.has(code)) errors.push(`Line ${i + 1}: ${code} is listed twice.`);
    seen.add(code);
    if (!en || en.length > 80) errors.push(`Line ${i + 1}: give an English name (at most 80 characters).`);
    if (pt.length > 80) errors.push(`Line ${i + 1}: the Portuguese name can be at most 80 characters.`);
    return { code, en, pt, active: r?.active !== false };
  });
  if (!clean.some((r) => r.active)) errors.push("Keep at least one entry switched on.");
  if (errors.length) return { ok: false, errors };
  db.transaction(() => {
    db.prepare(`UPDATE failure_codes SET active = 0, updated_at = datetime('now'), updated_by = ? WHERE kind = ? AND code NOT IN (${clean.map(() => "?").join(",")})`)
      .run(by || null, kind, ...clean.map((r) => r.code));
    const up = db.prepare(`
      INSERT INTO failure_codes (kind, code, label_en, label_pt, active, sort, updated_at, updated_by)
      VALUES (?, ?, ?, ?, ?, ?, datetime('now'), ?)
      ON CONFLICT(kind, code) DO UPDATE SET label_en = excluded.label_en, label_pt = excluded.label_pt, active = excluded.active,
        sort = excluded.sort, updated_at = excluded.updated_at, updated_by = excluded.updated_by
    `);
    clean.forEach((r, i) => up.run(kind, r.code, r.en, r.pt || null, r.active ? 1 : 0, (i + 1) * 10, by || null));
  })();
  return { ok: true, list: listFailureCodes(db, { all: true })[kind] };
}

/** Checks codes sent for a breakdown: each must be on its list (or already on the breakdown). */
export function readFailureCodes(db, body = {}, current = {}) {
  ensureFailureCodeTables(db);
  const values = {};
  const errors = [];
  for (const kind of Object.keys(FAILURE_KINDS)) {
    if (!Object.prototype.hasOwnProperty.call(body, kind)) continue;
    const code = String(body[kind] ?? "").trim().toUpperCase();
    if (!code) {
      values[`fc_${kind}`] = null;
      continue;
    }
    const known = db.prepare(`SELECT active FROM failure_codes WHERE kind = ? AND code = ?`).get(kind, code);
    if (!known || (!known.active && current[`fc_${kind}`] !== code)) errors.push(`${FAILURE_KINDS[kind]}: ${code} is not on the list.`);
    else values[`fc_${kind}`] = code;
  }
  return errors.length ? { ok: false, errors } : { ok: true, values };
}

/**
 * The foreman's codes for a breakdown. Confirmed (by, at) once all four are set; cleared again
 * when one is emptied. Returns the breakdown's codes.
 */
export function setBreakdownFailureCodes(db, breakdownId, values, by = "") {
  ensureFailureCodeTables(db);
  const cols = Object.keys(values);
  if (cols.length) db.prepare(`UPDATE breakdowns SET ${cols.map((c) => `${c} = ?`).join(", ")} WHERE id = ?`).run(...cols.map((c) => values[c]), breakdownId);
  const b = db.prepare(`SELECT fc_component, fc_problem, fc_cause, fc_action, fc_confirmed_at FROM breakdowns WHERE id = ?`).get(breakdownId);
  const complete = b && b.fc_component && b.fc_problem && b.fc_cause && b.fc_action;
  if (complete) db.prepare(`UPDATE breakdowns SET fc_confirmed_by = ?, fc_confirmed_at = datetime('now') WHERE id = ?`).run(by || null, breakdownId);
  else db.prepare(`UPDATE breakdowns SET fc_confirmed_by = NULL, fc_confirmed_at = NULL WHERE id = ?`).run(breakdownId);
  return breakdownFailureCodes(db, breakdownId);
}

/** The technician's suggestion (component, problem) from the portal; never over the foreman's confirmed codes. */
export function suggestBreakdownFailureCodes(db, breakdownId, { component, problem } = {}, by = "") {
  ensureFailureCodeTables(db);
  const b = db.prepare(`SELECT id, fc_confirmed_at FROM breakdowns WHERE id = ?`).get(breakdownId);
  if (!b || b.fc_confirmed_at) return false;
  const r = readFailureCodes(db, { ...(component ? { component } : {}), ...(problem ? { problem } : {}) });
  if (!r.ok || !Object.keys(r.values).length) return false;
  const cols = Object.keys(r.values);
  db.prepare(`UPDATE breakdowns SET ${cols.map((c) => `${c} = ?`).join(", ")}, fc_suggested_by = ?, fc_suggested_at = datetime('now') WHERE id = ?`)
    .run(...cols.map((c) => r.values[c]), by || null, breakdownId);
  return true;
}

/** { component, problem, cause, action, suggested_by, suggested_at, confirmed_by, confirmed_at } of one breakdown. */
export function breakdownFailureCodes(db, breakdownId) {
  ensureFailureCodeTables(db);
  const b = db.prepare(`SELECT ${COLS.join(", ")} FROM breakdowns WHERE id = ?`).get(breakdownId);
  if (!b) return null;
  return {
    component: b.fc_component, problem: b.fc_problem, cause: b.fc_cause, action: b.fc_action,
    suggested_by: b.fc_suggested_by, suggested_at: b.fc_suggested_at, confirmed_by: b.fc_confirmed_by, confirmed_at: b.fc_confirmed_at,
  };
}

/** The component code whose English or Portuguese name matches the "what failed" text (case-insensitive), or null. */
export function componentFromText(db, text) {
  const t = String(text || "").trim().toLowerCase();
  if (!t) return null;
  ensureFailureCodeTables(db);
  const r = db.prepare(`
    SELECT code FROM failure_codes
    WHERE kind = 'component' AND active = 1 AND (LOWER(label_en) = ? OR LOWER(COALESCE(label_pt, '')) = ? OR LOWER(code) = ?)
    LIMIT 1
  `).get(t, t, t);
  return r?.code || null;
}

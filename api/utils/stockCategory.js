// IRONLOG/api/utils/stockCategory.js — one category per store item.
// Stores stock is split into oils, G.E.T, parts, components and tyres. The
// category lives on parts.stock_category; every report and screen reads it
// instead of guessing from the item name. New items get an automatic
// suggestion (stock_category_source = 'auto') that stores can override
// ('manual'); automatic rows are re-evaluated when the rules improve.

export const STOCK_CATEGORIES = [
  { key: "oil", label: "Oils & lubricants" },
  { key: "get", label: "G.E.T" },
  { key: "part", label: "Parts" },
  { key: "component", label: "Components" },
  { key: "tyre", label: "Tyres" },
];
export const STOCK_CATEGORY_KEYS = STOCK_CATEGORIES.map((c) => c.key);
const LABELS = new Map(STOCK_CATEGORIES.map((c) => [c.key, c.label]));

export function stockCategoryLabel(key) {
  return LABELS.get(normalizeStockCategory(key)) || LABELS.get("part");
}

export function normalizeStockCategory(value) {
  const v = String(value || "").trim().toLowerCase().replace(/[^a-z]/g, "");
  if (!v) return null;
  if (["oil", "oils", "lube", "lubes", "lubricant", "lubricants", "oilslubricants", "grease", "coolant", "hydraulicoil"].includes(v)) return "oil";
  if (["get", "groundengagingtools", "groundengaging"].includes(v)) return "get";
  if (["component", "components", "majorcomponent", "majorcomponents"].includes(v)) return "component";
  if (["tyre", "tyres", "tire", "tires"].includes(v)) return "tyre";
  if (["part", "parts", "spare", "spares", "workshopspares"].includes(v)) return "part";
  return null;
}

const OIL_KINDS = new Set(["oil", "lube", "lubricant", "hydraulic_oil", "coolant", "grease", "hyd fluid", "hydraulic fluid"]);

// Words that make an item a small part even when it mentions a big assembly
// ("seal kit steering cylinder", "oil filter", "hub seal").
const PART_WORDS = /\b(seals?|kits?|o[\s-]?rings?|gaskets?|bearings?|bush(es|ing)?|filters?|bolts?|nuts?|washers?|hoses?|fittings?|pins?|shims?|sleeves?|plugs?|coils?|sensors?|switch(es)?|clamps?|brackets?|screws?|studs?|springs?|belts?|pads?|shoes?|mounts?|mountings?|cups?|cones?|collars?|flanges?|couplings?|glands?|valves? (spool|cap)|caps?|covers?|dryers?|fuses?|relays?|lamps?|bulbs?|wipers?|mirrors?|keys?|locks?|chains?|brushes?|diaphragms?|elements?)\b/;

const OIL_PATTERNS = [
  /\bfuchs\b/, /\brenolin\b/, /\btitan\b/, /\blupex\b/, /\bfricofin\b/, /\bshell\s+(rimula|tellus|spirax|gadus)\b/, /\bcastrol\b/, /\bmobil\b/, /\btotal\s+(rubia|azolla)\b/,
  /\b\d{1,2}w-?\d{2}\b/, /\bsae\s?-?\d{2,3}\b/, /\batf\b/, /\butto\b/, /\bto-?4\b/, /\bhlp\b/, /\blhm\+?\b/, /\bho\s?\d{2}\b/, /\bep\s?[0-3]\b/,
  /\b(engine|gear|hydraulic|transmission|axle|compressor|motor)\s+oil\b/, /\boil\s+(sae|\d)/, /\bgrease\b/, /\bcoolant\b/, /\banti-?freeze\b/, /\blubricants?\b/, /\blube\b/,
];

const TYRE_PATTERNS = [/\btyres?\b/, /\btires?\b/, /\b\d{2}\.\d{1,2}\s?-?r\s?\d{2}\b/, /\b\d{3}\/\d{2}\s?-?z?r\s?\d{2}(\.\d)?\b/, /\b\d{2}\.\d-\d{2}\b/, /\b\d{1,2}\.\d{2}\s?-?r?\s?\d{2}\b.*\b(ply|tl|tt)\b/];

const GET_PATTERNS = [
  /\bg\.?\s?e\.?\s?t\.?\b/, /ground engaging/, /\btips?\b/, /\bteeth\b/, /\btooth\b/, /cutting edges?/, /\bend ?bits?\b/,
  /side ?cutters?/, /\bshrouds?\b/, /\bripper (tip|shank|point)s?\b/, /\bbucket (adapter|lip|edge|pin)s?\b/,
  /\b[jkr]\d{3}\b.*\b(adapter|pin|retainer|lock)s?\b/, /\b(adapter|pin|retainer|lock)s?\b.*\b[jk]\d{3}\b/, /\bwear (plate|strip|bar)s?\b/, /\bgrader blades?\b/,
];

const COMPONENT_PATTERNS = [
  /\bengines?\b/, /\btransmissions?\b/, /\bgear ?box(es)?\b/, /\bfinal drives?\b/, /\bdifferentials?\b/, /\btorque converters?\b/,
  /\b(hydraulic|main|gear|piston|steering|water|fuel injection|transfer|charge|fan) pumps?\b/, /\bpumps?\b/,
  /\bcylinders?\b/, /\bcyl\b/, /\brams?\b/, /\bhubs?\b/, /\baxles?\b/, /\bturbo(charger)?s?\b/, /\balternators?\b/,
  /\bstarters?( motors?)?\b/, /\bradiators?\b/, /\b(hydraulic|swing|travel|drive|fan) motors?\b/, /\b(control|main|directional) valves?\b/,
  /\bvalves? (bank|block|directional)\b/, /\bswing (bearing|ring|gear)s?\b/, /\bslew(ing)? rings?\b/, /\bcompressors?\b/, /\bprop ?shafts?\b/,
];

/** Automatic category for a store item, from its kind, code and description. */
export function classifyStockItem({ part_code, part_name, consumable_kind } = {}) {
  const kind = String(consumable_kind || "").trim().toLowerCase();
  if (OIL_KINDS.has(kind)) return "oil";
  const code = String(part_code || "").trim().toLowerCase();
  // "less coupling" / "without pulley" describe what is NOT included.
  const name = String(part_name || "").trim().toLowerCase().replace(/[_]+/g, " ").replace(/\b(less|without|excl\.?|excluding)\s+\S+/g, " ");
  const text = `${name} ${code}`;
  if (/^mlf/.test(code)) return "oil"; // Fuchs lubricant stock codes
  // "Oil" only counts as a word, and not for oil seals, filters, coolers etc.
  const oilWord = /\boils?\b/.test(name) && !/\boil (seals?|filters?|coolers?|pumps?|caps?|pans?|lines?|pipes?|hoses?|pressure|level|sensors?)\b|\b(seals?|filters?) oil\b/.test(name);
  if (!PART_WORDS.test(name) && (oilWord || OIL_PATTERNS.some((re) => re.test(text)))) return "oil";
  if (TYRE_PATTERNS.some((re) => re.test(text)) && !/\b(valves?|gauges?|levers?|irons?|chains?|pressure)\b/.test(name)) return "tyre";
  if (GET_PATTERNS.some((re) => re.test(name))) return "get";
  if (/\bservice kits?\b|\b\d{2,5}\s?h(r|rs|our)?s?\b.*\bkits?\b/.test(name)) return "part";
  if (!PART_WORDS.test(name) && COMPONENT_PATTERNS.some((re) => re.test(name))) return "component";
  return "part";
}

function hasColumn(db, table, col) {
  try {
    return db.prepare(`PRAGMA table_info(${table})`).all().some((c) => c.name === col);
  } catch {
    return false;
  }
}

const prepared = new WeakSet();

/**
 * Adds parts.stock_category / stock_category_source and fills every item
 * that has no category yet or only an automatic one. Manual choices stay.
 * Runs once per database connection unless { refresh: true } is passed.
 */
export function ensureStockCategorySchema(db, { refresh = false } = {}) {
  if (!refresh && prepared.has(db)) return { updated: 0 };
  if (!hasColumn(db, "parts", "id")) return { updated: 0 };
  prepared.add(db);
  if (!hasColumn(db, "parts", "stock_category")) db.exec(`ALTER TABLE parts ADD COLUMN stock_category TEXT`);
  if (!hasColumn(db, "parts", "stock_category_source")) db.exec(`ALTER TABLE parts ADD COLUMN stock_category_source TEXT`);
  const kindCol = hasColumn(db, "parts", "consumable_kind") ? "consumable_kind" : "NULL AS consumable_kind";
  const rows = db.prepare(`
    SELECT id, part_code, part_name, ${kindCol}, stock_category, stock_category_source
    FROM parts
    WHERE stock_category IS NULL OR TRIM(stock_category) = '' OR COALESCE(stock_category_source, 'auto') = 'auto'
  `).all();
  const update = db.prepare(`UPDATE parts SET stock_category = ?, stock_category_source = 'auto' WHERE id = ?`);
  let updated = 0;
  db.transaction(() => {
    for (const r of rows) {
      const next = classifyStockItem(r);
      if (next !== r.stock_category || r.stock_category_source !== "auto") {
        update.run(next, r.id);
        updated++;
      }
    }
  })();
  return { updated };
}

/** Gives a newly created item its automatic category. */
export function autoCategorizePart(db, partId) {
  if (!hasColumn(db, "parts", "stock_category")) return;
  const kindCol = hasColumn(db, "parts", "consumable_kind") ? "consumable_kind" : "NULL AS consumable_kind";
  const row = db.prepare(`SELECT id, part_code, part_name, ${kindCol}, stock_category FROM parts WHERE id = ?`).get(Number(partId));
  if (!row || (row.stock_category && String(row.stock_category).trim())) return;
  db.prepare(`UPDATE parts SET stock_category = ?, stock_category_source = 'auto' WHERE id = ?`).run(classifyStockItem(row), row.id);
}

/** SQL expression for an item's category (defaults to 'part'). */
export function stockCategorySql(alias = "p") {
  return `COALESCE(NULLIF(TRIM(${alias}.stock_category), ''), 'part')`;
}

/** SQL condition: the item is an oil / lubricant. */
export function oilPartSql(alias = "p") {
  return `(${stockCategorySql(alias)} = 'oil')`;
}

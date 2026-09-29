// IRONLOG/api/utils/translate.js — Portuguese ⇄ English for workshop text
// (operator fault comments, inspection notes). Uses whichever AI provider is
// configured for BORRIS; returns null when none is set up so callers can fall
// back to the checklist glossary.

const LANGUAGES = { en: "English", pt: "Portuguese as spoken in Mozambique" };

export function translationMessages(text, to = "en") {
  const target = LANGUAGES[to] || LANGUAGES.en;
  return [
    {
      role: "system",
      content: [
        `Translate the user's text into ${target}.`,
        "It is written by operators, artisans or foremen at a quarry and mining workshop: machine parts, faults, pre-start checks.",
        "Use normal workshop terms. Keep asset codes, part numbers, numbers and units exactly as written.",
        "If the text is already in the target language, return it unchanged.",
        "Reply with the translation only: no quotes, notes or explanations.",
      ].join(" "),
    },
    { role: "user", content: String(text || "") },
  ];
}

/** Strips wrapping quotes and "Translation:" style prefixes some models add. */
export function cleanTranslation(out) {
  let s = String(out || "").trim();
  s = s.replace(/^(translation|tradução)\s*:\s*/i, "");
  if (/^(["'“]).*(["'”])$/s.test(s)) s = s.slice(1, -1).trim();
  return s;
}

/**
 * chat(messages, opts) returns reply text or null (see aiChatText in ironmind.js).
 * Resolves to the translation, "" for empty input, or null when unavailable.
 */
export async function translateText(text, { to = "en", chat } = {}) {
  const src = String(text || "").trim();
  if (!src) return "";
  if (typeof chat !== "function") return null;
  const out = await chat(translationMessages(src, to), {
    temperature: 0,
    max_tokens: Math.min(1200, 60 + src.length * 2),
    timeout_ms: 20000,
  });
  if (out == null) return null;
  const cleaned = cleanTranslation(out);
  return cleaned || null;
}

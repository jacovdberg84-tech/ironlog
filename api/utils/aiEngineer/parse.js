export function extractJsonObject(text) {
  const s = String(text || "").trim();
  const fenced = s.match(/```(?:json)?\s*([\s\S]*?)```/i);
  const candidate = fenced?.[1]?.trim() || s;
  try {
    return JSON.parse(candidate);
  } catch {
    const start = candidate.indexOf("{");
    const end = candidate.lastIndexOf("}");
    if (start >= 0 && end > start) {
      try {
        return JSON.parse(candidate.slice(start, end + 1));
      } catch {
        return null;
      }
    }
    return null;
  }
}

export function extractUnifiedDiff(text) {
  const s = String(text || "");
  const fenced = s.match(/```diff\s*([\s\S]*?)```/i);
  if (fenced?.[1]) return fenced[1].trim();
  const fallback = s.match(/(diff --git[\s\S]*)/i);
  return fallback?.[1]?.trim() || "";
}

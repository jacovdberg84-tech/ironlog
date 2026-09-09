export function isManualQuestion(question) {
 return /\b(manuals?|bulletins?|fault codes?|torque|specifications?|part\s*(?:numbers?|no\.?|#)|parts?\s+(?:manual|catalog(?:ue)?|number|lookup)|file\s+(?:location|path)|hub seals?)\b/i.test(String(question));
}
export function assetCandidates(question) {
 return [...new Set(String(question).toUpperCase().match(/\b(?=[A-Z0-9-]*[A-Z])(?=[A-Z0-9-]*\d)[A-Z][A-Z0-9-]{2,}\b/g)||[])];
}

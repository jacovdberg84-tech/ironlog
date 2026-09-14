const YMD_RE = /^(\d{4})-(\d{2})-(\d{2})$/;

function parseYmd(value, label) {
  const text = String(value || "").trim();
  const match = text.match(YMD_RE);
  if (!match) throw new Error(`${label} must be a valid YYYY-MM-DD date`);

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    throw new Error(`${label} must be a valid calendar date`);
  }
  return { text, date };
}

function localYmd(now) {
  const date = new Date(now);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

/**
 * Validated selected-date range for a FAMS catch-up import.
 * FAMS expects slash-delimited dates, while IRONLOG stores ISO calendar dates.
 */
export function famsSelectedDateRange({ startDate, endDate, now = new Date(), maxDays = 366 } = {}) {
  const start = parseYmd(startDate, "From date");
  const end = parseYmd(endDate, "To date");
  const today = parseYmd(localYmd(now), "Today");

  if (start.date > end.date) throw new Error("From date cannot be after To date");
  if (end.date > today.date) throw new Error("To date cannot be later than today");

  const inclusiveDays = Math.floor((end.date - start.date) / 86_400_000) + 1;
  if (inclusiveDays > maxDays) throw new Error(`Select a period of ${maxDays} days or less`);

  return {
    startYmd: start.text,
    endYmd: end.text,
    startSlash: start.text.replaceAll("-", "/"),
    endSlash: end.text.replaceAll("-", "/"),
  };
}

// IRONLOG/api/utils/companyProfile.js — the customer company IronLog runs for.
//
// IronLog is the product; each customer (AML today) has its own name, short
// name for report headers, contact details, currency, time zone, country and
// site location. These live in report_settings (keys company_*), next to the
// PDF logo and company/site names that already sit there, and are edited by an
// administrator under Admin → Company profile. Nothing customer-specific is
// written into report code any more.
//
// Existing installs keep their wording: the first time this runs on a database
// that already holds data, the profile is filled with the values that used to
// be fixed in the code (AML, Palma, Mozambique). A new, empty database starts
// neutral and is set up by its administrator.
import { ensureReportSettingsSchema, getPdfReportBranding, getReportSetting, setReportSetting } from "./reportSettings.js";

export const COMPANY_FIELDS = {
  company_name: { label: "Company name", max: 120 },
  short_name: { label: "Short name on report headers", max: 24 },
  site_name: { label: "Site name on reports", max: 120 },
  address: { label: "Address", max: 300 },
  phone: { label: "Phone", max: 60 },
  email: { label: "Email", max: 120 },
  website: { label: "Website", max: 120 },
  registration_no: { label: "Company registration number", max: 60 },
  currency: { label: "Currency", max: 3 },
  timezone: { label: "Time zone", max: 60 },
  country: { label: "Country", max: 60 },
  location: { label: "Site location (for weather and security news)", max: 120 },
  latitude: { label: "Latitude", max: 20 },
  longitude: { label: "Longitude", max: 20 },
  nearby: { label: "Nearby towns to watch for security news (comma separated)", max: 400 },
};

// What existing installs showed before these became settings. The site name is
// left out on purpose: it is shared with every PDF header, which showed none.
export const LEGACY_PROFILE = {
  short_name: "AML",
  currency: "USD",
  timezone: "Africa/Maputo",
  country: "Mozambique",
  location: "Palma, Cabo Delgado",
  latitude: "-10.7768",
  longitude: "40.4764",
  nearby: "Palma, Afungi, Mocímboa da Praia, Nangade, Muidumbe, Mueda, Macomia, Quionga",
};

// A new company starts in the time zone its server was set up with (IRONLOG_TIMEZONE), else UTC.
const neutralProfile = () => ({ currency: "USD", timezone: process.env.IRONLOG_TIMEZONE || "UTC" });
const VERSION_KEY = "company_profile_version";

const keyOf = (field) => (field === "company_name" || field === "site_name" ? `pdf_${field}` : `company_${field}`);

function hasData(db) {
  try {
    return Number(db.prepare(`SELECT COUNT(*) AS n FROM assets`).get()?.n || 0) > 0;
  } catch {
    return false;
  }
}

/** Fills the profile once: legacy values on an existing install, neutral ones on a new database. */
export function ensureCompanyProfile(db) {
  ensureReportSettingsSchema(db);
  if (getReportSetting(db, VERSION_KEY)) return;
  const seed = hasData(db) ? LEGACY_PROFILE : neutralProfile();
  for (const [field, value] of Object.entries(seed)) {
    if (!getReportSetting(db, keyOf(field))) setReportSetting(db, keyOf(field), value);
  }
  setReportSetting(db, VERSION_KEY, "1");
}

function clean(field, value) {
  const s = String(value ?? "").trim();
  // A currency is checked whole, so "dollars" is refused rather than cut to "DOL".
  if (field === "currency") return s.toUpperCase();
  return s.slice(0, COMPANY_FIELDS[field].max);
}

// ISO 4217 codes the runtime knows (USD, ZAR, MZN…), so "DOL" is refused.
const CURRENCIES = new Set(typeof Intl.supportedValuesOf === "function" ? Intl.supportedValuesOf("currency") : []);
function validCurrency(code) {
  return /^[A-Z]{3}$/.test(code) && (!CURRENCIES.size || CURRENCIES.has(code));
}

function validTimeZone(tz) {
  try {
    new Intl.DateTimeFormat("en-CA", { timeZone: tz });
    return true;
  } catch {
    return false;
  }
}

/** The whole profile, with header prefix and logo details. */
export function getCompanyProfile(db) {
  ensureCompanyProfile(db);
  const out = {};
  for (const field of Object.keys(COMPANY_FIELDS)) out[field] = getReportSetting(db, keyOf(field));
  // The PDF settings may name the company by code only.
  const branding = getPdfReportBranding(db);
  if (!out.company_name) out.company_name = branding.company_name || "";
  if (!out.site_name) out.site_name = branding.site_name || "";
  if (!out.currency) out.currency = "USD";
  if (!out.timezone || !validTimeZone(out.timezone)) out.timezone = "UTC";
  out.logo_url = branding.company_logo_url || "";
  out.product_name = "IronLog";
  out.header_prefix = out.short_name || out.company_name || "";
  return out;
}

/** Saves the fields given; returns the errors found (nothing is saved when there are any). */
export function saveCompanyProfile(db, body = {}) {
  ensureCompanyProfile(db);
  const next = {};
  for (const field of Object.keys(COMPANY_FIELDS)) {
    if (Object.prototype.hasOwnProperty.call(body, field)) next[field] = clean(field, body[field]);
  }
  const errors = [];
  if ("currency" in next && next.currency && !validCurrency(next.currency)) errors.push("Currency must be a 3-letter code such as USD or ZAR.");
  if ("timezone" in next && next.timezone && !validTimeZone(next.timezone)) errors.push("Time zone not recognised. Use a name such as Africa/Maputo or Africa/Johannesburg.");
  for (const f of ["latitude", "longitude"]) {
    if (f in next && next[f] !== "") {
      const n = Number(next[f]);
      const limit = f === "latitude" ? 90 : 180;
      if (!Number.isFinite(n) || Math.abs(n) > limit) errors.push(`${COMPANY_FIELDS[f].label} must be a number between -${limit} and ${limit}.`);
    }
  }
  if ("email" in next && next.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(next.email)) errors.push("Email address does not look right.");
  if (errors.length) return { ok: false, errors };
  for (const [field, value] of Object.entries(next)) setReportSetting(db, keyOf(field), value);
  return { ok: true, profile: getCompanyProfile(db) };
}

/** "AML · IronLog — Daily operations report", or "IronLog — …" with no company set. */
export function reportTitle(profile, title) {
  return profile.header_prefix ? `${profile.header_prefix} · IronLog — ${title}` : `IronLog — ${title}`;
}

/** "AML / DAILY OPERATIONS" style banner, or just the words with no company set. */
export function reportBanner(profile, words) {
  const upper = String(words).toUpperCase();
  return profile.header_prefix ? `${profile.header_prefix.toUpperCase()} / ${upper}` : upper;
}

/** The site June's daily briefing describes, from the profile. */
export function briefingSiteFromProfile(profile) {
  const num = (v) => (String(v ?? "").trim() !== "" && Number.isFinite(Number(v)) ? Number(v) : null);
  return {
    name: profile.location || profile.site_name || profile.company_name || "",
    country: profile.country || "",
    latitude: num(profile.latitude),
    longitude: num(profile.longitude),
    timezone: profile.timezone || "UTC",
    nearby: String(profile.nearby || "").split(",").map((s) => s.trim()).filter(Boolean),
  };
}

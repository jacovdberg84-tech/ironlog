import { db } from "../db/client.js";
import { buildPdfBuffer, sectionTitle, kvGrid, ensurePageSpace } from "../utils/pdfGenerator.js";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import fs from "fs";
import path from "path";

const LANGS = new Set(["en", "af", "zu", "pt"]);
const DOC_TYPES = new Set(["SOP", "Site Instruction", "Method Statement", "Checklist", "Risk Note"]);

function nowIso() {
  return new Date().toISOString();
}

function getUser(req) {
  return String(req.headers["x-user-name"] || "session-user").trim() || "session-user";
}

function hasColumn(tableName, columnName) {
  const rows = db.prepare(`PRAGMA table_info(${tableName})`).all();
  return rows.some((r) => String(r.name || "").toLowerCase() === String(columnName || "").toLowerCase());
}

function ensureColumn(tableName, columnSqlDef, columnName) {
  if (!hasColumn(tableName, columnName)) {
    db.prepare(`ALTER TABLE ${tableName} ADD COLUMN ${columnSqlDef}`).run();
  }
}

function toBool(v) {
  const s = String(v || "").trim().toLowerCase();
  return s === "1" || s === "true" || s === "yes" || s === "y";
}

function safeFileName(v, fallback = "document") {
  return String(v || fallback)
    .replace(/[^\w.-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 80) || fallback;
}

function stripHeaderFromDraftText(fullText) {
  const source = String(fullText || "-");
  return source.replace(
    /^Document Type:[\s\S]*?Date:\s*[0-9]{4}-[0-9]{2}-[0-9]{2}\s*/i,
    ""
  ).trim() || source;
}

function splitLines(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .split("\n");
}

function resolveHeaderByInputOrLatest(headerId) {
  const id = Number(headerId || 0);
  if (id > 0) {
    return db.prepare(`
      SELECT id, name, site_name, department, prepared_by, approved_by, revision
      FROM doc_headers
      WHERE id = ?
    `).get(id);
  }
  return db.prepare(`
    SELECT id, name, site_name, department, prepared_by, approved_by, revision
    FROM doc_headers
    ORDER BY id DESC
    LIMIT 1
  `).get();
}

function rowsForApprovedRegister(currentOnly) {
  const where = currentOnly
    ? `
      WHERE d.decision = 'approved'
        AND NOT EXISTS (
          SELECT 1
          FROM doc_drafts nx
          WHERE nx.supersedes_draft_id = d.id
            AND nx.decision = 'approved'
        )
    `
    : `WHERE d.decision = 'approved'`;
  return db.prepare(`
      SELECT
        d.id,
        d.doc_type,
        d.title,
        d.language,
        d.revision_no,
        d.header_id,
        h.name AS header_name,
        h.department,
        h.site_name,
        d.decided_by,
        d.decided_at,
        d.created_at
      FROM doc_drafts d
      LEFT JOIN doc_headers h ON h.id = d.header_id
      ${where}
      ORDER BY d.doc_type ASC, d.title ASC, COALESCE(d.revision_no, 0) DESC, d.id DESC
      LIMIT 2000
    `).all();
}

function normalizeLang(v) {
  const lang = String(v || "en").trim().toLowerCase();
  return LANGS.has(lang) ? lang : "en";
}

function normalizeDocType(v) {
  const t = String(v || "").trim();
  return DOC_TYPES.has(t) ? t : "SOP";
}

function buildDraftHeaderText({ docType, title, header }) {
  const stamp = nowIso().slice(0, 10);
  return [
    `Document Type: ${docType}`,
    `Title: ${title}`,
    `Site: ${header.site_name || "-"}`,
    `Department: ${header.department || "-"}`,
    `Prepared By: ${header.prepared_by || "-"}`,
    `Approved By: ${header.approved_by || "-"}`,
    `Revision: ${header.revision || "-"}`,
    `Date: ${stamp}`,
  ].join("\n");
}

function getAiConfig(preferredProvider = "") {
  const openaiKey = process.env.OPENAI_API_KEY;
  const azureKey = process.env.AZURE_OPENAI_API_KEY;
  const openaiModel = process.env.OPENAI_MODEL || "gpt-4o-mini";

  const azureEndpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const azureDeployment = process.env.AZURE_OPENAI_DEPLOYMENT;
  const azureApiVersion = process.env.AZURE_OPENAI_API_VERSION || "2024-02-15-preview";
  const foundryKey =
    process.env.FOUNDRY_API_KEY ||
    process.env.AZURE_FOUNDRY_API_KEY ||
    process.env.AZURE_OPENAI_API_KEY ||
    "";
  const foundryEndpoint =
    process.env.FOUNDRY_ENDPOINT ||
    process.env.AZURE_FOUNDRY_ENDPOINT ||
    process.env.AZURE_EXISTING_AIPROJECT_ENDPOINT ||
    "";
  const foundryModel =
    process.env.FOUNDRY_MODEL ||
    process.env.AZURE_FOUNDRY_MODEL ||
    process.env.OPENAI_MODEL ||
    "gpt-4o-mini";
  const requested = String(preferredProvider || process.env.AI_PROVIDER || "").trim().toLowerCase();
  const hasFoundry = Boolean(foundryKey && foundryEndpoint);

  if (requested === "foundry") {
    if (hasFoundry) {
      return {
        provider: "foundry",
        apiKey: foundryKey,
        endpoint: foundryEndpoint,
        model: foundryModel,
      };
    }
  }

  if (openaiKey) {
    return { provider: "openai", apiKey: openaiKey, model: openaiModel };
  }
  if (azureKey && azureEndpoint && azureDeployment) {
    return {
      provider: "azure_openai",
      apiKey: azureKey,
      endpoint: azureEndpoint,
      deployment: azureDeployment,
      apiVersion: azureApiVersion,
    };
  }
  if (hasFoundry) {
    return {
      provider: "foundry",
      apiKey: foundryKey,
      endpoint: foundryEndpoint,
      model: foundryModel,
    };
  }
  return { provider: null };
}

function normalizeFoundryChatEndpoint(endpoint) {
  const base = String(endpoint || "").trim().replace(/\/+$/, "");
  if (!base) return "";
  if (/\/openai\/v1$/i.test(base)) return `${base}/chat/completions`;
  if (/\/api\/projects\//i.test(base)) return `${base}/openai/v1/chat/completions`;
  return `${base}/openai/v1/chat/completions`;
}

function findRelevantLegalDocs(query, limit = 5) {
  const q = String(query || "").trim().toLowerCase();
  if (!q) return [];
  const tokens = q.split(/\s+/).filter((t) => t.length >= 3).slice(0, 8);
  if (!tokens.length) return [];

  const all = db.prepare(`
    SELECT id, department, title, doc_type, owner, status, created_at
    FROM legal_documents
    WHERE active = 1
    ORDER BY id DESC
    LIMIT 300
  `).all();

  const scored = all
    .map((r) => {
      const hay = `${r.title || ""} ${r.doc_type || ""} ${r.department || ""} ${r.owner || ""}`.toLowerCase();
      let score = 0;
      for (const t of tokens) if (hay.includes(t)) score += 1;
      if (score <= 0) return null;
      return { ...r, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score || b.id - a.id)
    .slice(0, limit);
  return scored;
}

function htmlToPlainText(html) {
  return String(html || "")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeSourceUrl(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  try {
    const u = new URL(s);
    if (u.hostname.toLowerCase().includes("duckduckgo.com")) {
      const redirected = u.searchParams.get("uddg");
      if (redirected && /^https?:\/\//i.test(redirected)) return redirected;
    }
    return u.toString();
  } catch {
    return "";
  }
}

function parseTrustedDomains(input) {
  return String(input || "")
    .split(",")
    .map((d) => d.trim().toLowerCase())
    .filter(Boolean)
    .map((d) => d.replace(/^\*\./, ""))
    .filter((d) => /^[a-z0-9.-]+$/.test(d));
}

function isTrustedUrl(url, trustedDomains) {
  if (!Array.isArray(trustedDomains) || !trustedDomains.length) return true;
  try {
    const host = String(new URL(url).hostname || "").toLowerCase();
    return trustedDomains.some((d) => host === d || host.endsWith(`.${d}`));
  } catch {
    return false;
  }
}

async function fetchWebSearchContext(query, limit = 3, trustedDomains = []) {
  const q = String(query || "").trim();
  if (!q) return [];
  try {
    const url = `https://duckduckgo.com/html/?q=${encodeURIComponent(q)}`;
    const res = await fetch(url, {
      headers: { "User-Agent": "IRONLOG/1.0" },
    });
    if (!res.ok) return [];
    const html = await res.text();
    const matches = Array.from(
      html.matchAll(/<a[^>]*class="[^"]*result__a[^"]*"[^>]*href="([^"]+)"[^>]*>[\s\S]*?<\/a>/gi)
    );
    const urls = matches
      .map((m) => normalizeSourceUrl(m[1]))
      .filter((u) => /^https?:\/\//i.test(u))
      .filter((u) => isTrustedUrl(u, trustedDomains))
      .slice(0, Math.max(1, limit));

    const out = [];
    for (const u of urls) {
      try {
        const page = await fetch(u, { headers: { "User-Agent": "IRONLOG/1.0" } });
        if (!page.ok) continue;
        const pageHtml = await page.text();
        const snippet = htmlToPlainText(pageHtml).slice(0, 1200);
        if (!snippet) continue;
        out.push({ url: u, snippet });
      } catch {
        // Ignore failing source and continue.
      }
    }
    return out.slice(0, limit);
  } catch {
    return [];
  }
}

async function generateAIDraftBody({ language, docType, scope, hazards, controls, extraNotes, preferFoundry = false }) {
  const cfg = getAiConfig(preferFoundry ? "foundry" : "");
  if (!cfg.provider) return null;
  const isChecklist = String(docType || "").toLowerCase() === "checklist";

  const langInstruction =
    language === "af"
      ? "Write the content in Afrikaans."
      : language === "zu"
        ? "Write the content in isiZulu."
        : language === "pt"
          ? "Write the content in Portuguese."
          : "Write the content in English.";

  const structureInstruction = isChecklist
    ? [
        "Generate the body of the document ONLY (no header).",
        "Use a practical checklist format with short checkbox-ready items.",
        "Suggested sections:",
        "1. Purpose",
        "2. Scope",
        "3. Pre-Start Checks",
        "4. During Task Checks",
        "5. Post-Task Close-Out",
        "6. Sign-Off",
      ].join("\n")
    : [
        "Generate the body of the document ONLY (no header). Use sections 1 to 8 with the following headings in the selected language:",
        "1. Purpose (or Doel/Inhloso/Objetivo)",
        "2. Scope (or Omvang/Ububanzi/Escopo)",
        "3. Responsibilities",
        "4. Hazards / Risks",
        "5. Controls",
        "6. Procedure",
        "7. Records and Evidence",
        "8. Revision Control",
      ].join("\n");

  const user = [
    `Document Type: ${docType}`,
    `Scope / objective: ${scope || "-"}`,
    `Hazards / risks: ${hazards || "-"}`,
    `Controls / PPE: ${controls || "-"}`,
    `Extra notes / requirements: ${extraNotes || "-"}`,
    "",
    structureInstruction,
    "",
    "Rules:",
    "- Keep it compliance-focused and practical.",
    "- Use short bullet points and clear section headings.",
    "- Do not include any meta commentary.",
    "- Do not include the literal labels 'User request', 'Related legal docs', or 'Web sources' in the output.",
    "- Do not copy prompt instructions into the final document.",
  ].join("\n");

  try {
    if (cfg.provider === "openai") {
      const res = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${cfg.apiKey}`,
        },
        body: JSON.stringify({
          model: cfg.model,
          temperature: Number(process.env.DOC_AI_TEMPERATURE ?? 0.3),
          max_tokens: Number(process.env.DOC_AI_MAX_TOKENS ?? 900),
          messages: [
            {
              role: "system",
              content:
                "You draft operational site compliance documents. Return only the requested document body.",
            },
            { role: "user", content: `${langInstruction}\n\n${user}` },
          ],
        }),
      });

      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content;
      return typeof text === "string" && text.trim() ? text.trim() : null;
    }

    // Azure OpenAI
    if (cfg.provider === "azure_openai") {
      const url = `${cfg.endpoint.replace(/\/$/, "")}/openai/deployments/${encodeURIComponent(
        cfg.deployment
      )}/chat/completions?api-version=${encodeURIComponent(cfg.apiVersion)}`;

      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": cfg.apiKey,
        },
        body: JSON.stringify({
          temperature: Number(process.env.DOC_AI_TEMPERATURE ?? 0.3),
          max_tokens: Number(process.env.DOC_AI_MAX_TOKENS ?? 900),
          messages: [
            {
              role: "system",
              content:
                "You draft operational site compliance documents. Return only the requested document body.",
            },
            { role: "user", content: `${langInstruction}\n\n${user}` },
          ],
        }),
      });

      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content;
      return typeof text === "string" && text.trim() ? text.trim() : null;
    }
    if (cfg.provider === "foundry") {
      const url = normalizeFoundryChatEndpoint(cfg.endpoint);
      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": cfg.apiKey,
          Authorization: `Bearer ${cfg.apiKey}`,
        },
        body: JSON.stringify({
          model: cfg.model,
          temperature: Number(process.env.DOC_AI_TEMPERATURE ?? 0.3),
          max_tokens: Number(process.env.DOC_AI_MAX_TOKENS ?? 900),
          messages: [
            {
              role: "system",
              content:
                "You draft operational site compliance documents. Return only the requested document body.",
            },
            { role: "user", content: `${langInstruction}\n\n${user}` },
          ],
        }),
      });
      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content || data?.output_text;
      return typeof text === "string" && text.trim() ? text.trim() : null;
    }
  } catch (e) {
    // Keep logs safe: do not print API keys.
    console.error("[docs ai] generate failed:", e?.message || e);
    return null;
  }

  return null;
}

async function generateAiDocumentReview({ content, docType, instructions, preferFoundry = false }) {
  const cfg = getAiConfig(preferFoundry ? "foundry" : "");
  if (!cfg.provider) return null;

  const cleanContent = String(content || "").trim();
  if (!cleanContent) return null;

  const requestUser = [
    `Document Type: ${docType || "General compliance document"}`,
    `Instructions: ${instructions || "Suggest corrections, improve clarity, and provide a clean final text output. Mark proposed changes clearly if possible."}`,
    "\nOriginal Content:\n",
    cleanContent,
  ].join("\n");

  const systemInstruction =
    "You are an expert compliance document editor. Review the text and suggest corrections. Output a finalized corrected document and a brief change summary.";

  try {
    if (cfg.provider === "openai") {
      const res = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${cfg.apiKey}`,
        },
        body: JSON.stringify({
          model: cfg.model,
          temperature: Number(process.env.DOC_AI_TEMPERATURE ?? 0.2),
          max_tokens: Number(process.env.DOC_AI_MAX_TOKENS ?? 1200),
          messages: [
            { role: "system", content: systemInstruction },
            { role: "user", content: requestUser },
          ],
        }),
      });

      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content;
      return typeof text === "string" && text.trim() ? text.trim() : null;
    }

    if (cfg.provider === "azure_openai") {
      const url = `${cfg.endpoint.replace(/\/$/, "")}/openai/deployments/${encodeURIComponent(
        cfg.deployment
      )}/chat/completions?api-version=${encodeURIComponent(cfg.apiVersion)}`;

      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": cfg.apiKey,
        },
        body: JSON.stringify({
          temperature: Number(process.env.DOC_AI_TEMPERATURE ?? 0.2),
          max_tokens: Number(process.env.DOC_AI_MAX_TOKENS ?? 1200),
          messages: [
            { role: "system", content: systemInstruction },
            { role: "user", content: requestUser },
          ],
        }),
      });

      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content;
      return typeof text === "string" && text.trim() ? text.trim() : null;
    }

    if (cfg.provider === "foundry") {
      const url = normalizeFoundryChatEndpoint(cfg.endpoint);
      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": cfg.apiKey,
          Authorization: `Bearer ${cfg.apiKey}`,
        },
        body: JSON.stringify({
          model: cfg.model,
          temperature: Number(process.env.DOC_AI_TEMPERATURE ?? 0.2),
          max_tokens: Number(process.env.DOC_AI_MAX_TOKENS ?? 1200),
          messages: [
            { role: "system", content: systemInstruction },
            { role: "user", content: requestUser },
          ],
        }),
      });

      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content || data?.output_text;
      return typeof text === "string" && text.trim() ? text.trim() : null;
    }
  } catch (e) {
    console.error("[docs ai] review failed:", e?.message || e);
    return null;
  }

  return null;
}

async function generateAiTechAnswer({ machine, problem, context = "", preferFoundry = false }) {
  const cfg = getAiConfig(preferFoundry ? "foundry" : "");
  if (!cfg.provider) return null;

  const requestUser = [
    `Machine/Asset: ${machine || "Unknown machine"}`,
    `Problem: ${problem || "No problem provided"}`,
    context ? `Context: ${context}` : "",
    "",
    "Return practical troubleshooting steps as a numbered list.",
    "Keep steps short and actionable.",
    "Start with safe checks, then fluid/electrical/mechanical checks, then escalation.",
    "Do not include markdown code blocks.",
  ].filter(Boolean).join("\n");

  const systemInstruction =
    "You are a heavy equipment diagnostic assistant for site mechanics. Provide concise, safe, practical troubleshooting steps. Use numbered steps only.";

  try {
    if (cfg.provider === "openai") {
      const res = await fetch("https://api.openai.com/v1/chat/completions", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${cfg.apiKey}`,
        },
        body: JSON.stringify({
          model: cfg.model,
          temperature: Number(process.env.DOC_AI_TEMPERATURE ?? 0.2),
          max_tokens: Number(process.env.DOC_AI_MAX_TOKENS ?? 900),
          messages: [
            { role: "system", content: systemInstruction },
            { role: "user", content: requestUser },
          ],
        }),
      });

      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content;
      return typeof text === "string" && text.trim() ? text.trim() : null;
    }

    if (cfg.provider === "azure_openai") {
      const url = `${cfg.endpoint.replace(/\/$/, "")}/openai/deployments/${encodeURIComponent(
        cfg.deployment
      )}/chat/completions?api-version=${encodeURIComponent(cfg.apiVersion)}`;

      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": cfg.apiKey,
        },
        body: JSON.stringify({
          temperature: Number(process.env.DOC_AI_TEMPERATURE ?? 0.2),
          max_tokens: Number(process.env.DOC_AI_MAX_TOKENS ?? 900),
          messages: [
            { role: "system", content: systemInstruction },
            { role: "user", content: requestUser },
          ],
        }),
      });

      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content;
      return typeof text === "string" && text.trim() ? text.trim() : null;
    }

    if (cfg.provider === "foundry") {
      const url = normalizeFoundryChatEndpoint(cfg.endpoint);
      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "api-key": cfg.apiKey,
          Authorization: `Bearer ${cfg.apiKey}`,
        },
        body: JSON.stringify({
          model: cfg.model,
          temperature: Number(process.env.DOC_AI_TEMPERATURE ?? 0.2),
          max_tokens: Number(process.env.DOC_AI_MAX_TOKENS ?? 900),
          messages: [
            { role: "system", content: systemInstruction },
            { role: "user", content: requestUser },
          ],
        }),
      });

      const data = await res.json();
      const text = data?.choices?.[0]?.message?.content || data?.output_text;
      return typeof text === "string" && text.trim() ? text.trim() : null;
    }
  } catch (e) {
    console.error("[docs ai] ask failed:", e?.message || e);
    return null;
  }

  return null;
}

function draftWithTemplate({
  language,
  docType,
  title,
  header,
  scope,
  hazards,
  controls,
  extraNotes,
}) {
  const top = buildDraftHeaderText({ docType, title, header });
  const isChecklist = String(docType || "").toLowerCase() === "checklist";

  if (isChecklist) {
    const checklistEn = `
1. Purpose
${scope || "Confirm readiness and safe execution before, during, and after work."}

2. Scope
Applies to relevant personnel and contractors performing this task.

3. Pre-Start Checks
- [ ] Permit and authorization confirmed.
- [ ] Toolbox talk completed and understood.
- [ ] PPE and isolation controls in place.
- [ ] Hazards reviewed: ${hazards || "site/task hazards identified"}.

4. During Task Checks
- [ ] Work follows approved method.
- [ ] Controls maintained: ${controls || "PPE, barricading, and lockout controls"}.
- [ ] Deviations stopped and reported immediately.

5. Post-Task Close-Out
- [ ] Work area cleaned and made safe.
- [ ] Equipment returned to normal safe condition.
- [ ] Findings and actions recorded.

6. Sign-Off
- Supervisor: ____________________  Date: __________
- Team Lead: _____________________  Date: __________

Additional Notes
${extraNotes || "-"}
`.trim();

    const checklistAf = `
1. Doel
${scope || "Bevestig gereedheid en veilige uitvoering voor, tydens en na werk."}

2. Omvang
Van toepassing op relevante personeel en kontrakteurs wat hierdie taak uitvoer.

3. Voor-Begin Kontroles
- [ ] Permit en magtiging bevestig.
- [ ] Toolbox-gesprek voltooi en verstaan.
- [ ] PPE en isolasiebeheer is in plek.
- [ ] Gevare hersien: ${hazards || "terrein-/taakgevare geidentifiseer"}.

4. Tydens Taak Kontroles
- [ ] Werk volg goedgekeurde metode.
- [ ] Beheermaatreels gehandhaaf: ${controls || "PPE, afbakening en lockout-beheer"}.
- [ ] Afwykings onmiddellik gestop en gerapporteer.

5. Na-Taak Afsluiting
- [ ] Werkarea skoongemaak en veilig gemaak.
- [ ] Toerusting terug na normale veilige toestand.
- [ ] Bevindings en aksies aangeteken.

6. Aftekening
- Toesighouer: ____________________  Datum: __________
- Spanleier: ______________________  Datum: __________

Bykomende Notas
${extraNotes || "-"}
`.trim();

    const checklistZu = `
1. Inhloso
${scope || "Qinisekisa ukulungela nokusebenza ngokuphepha ngaphambi, ngesikhathi nangemuva komsebenzi."}

2. Ububanzi
Kusebenza kubasebenzi nakonkontileka abafanele abenza lo msebenzi.

3. Ukuhlola Ngaphambi Kokuqala
- [ ] Imvume negunya kuqinisekisiwe.
- [ ] Toolbox talk yenziwe futhi yaqondwa.
- [ ] I-PPE nezilawuli zokuhlukanisa zikho.
- [ ] Izingozi zibuyekeziwe: ${hazards || "izingozi zesiza/umsebenzi zikhonjwe"}.

4. Ukuhlola Ngesikhathi Somsebenzi
- [ ] Umsebenzi ulandela indlela evunyelwe.
- [ ] Izilawuli zigcinwa: ${controls || "PPE, ukwahlukanisa indawo, ne-lockout controls"}.
- [ ] Ukuphambuka kumiswa futhi kubikwe ngokushesha.

5. Ukuvala Ngemuva Komsebenzi
- [ ] Indawo yomsebenzi ihlanzekile futhi iphephile.
- [ ] Imishini ibuyiselwe esimweni esiphephile.
- [ ] Okutholakele nezinyathelo kuqoshiwe.

6. Ukusayina
- Umphathi: ____________________  Usuku: __________
- Umholi wethimba: _____________  Usuku: __________

Amanothi Engeziwe
${extraNotes || "-"}
`.trim();

    const checklistPt = `
1. Objetivo
${scope || "Confirmar prontidao e execucao segura antes, durante e apos o trabalho."}

2. Escopo
Aplica-se ao pessoal e contratados relevantes que executam esta tarefa.

3. Verificacoes Pre-Inicio
- [ ] Permissao e autorizacao confirmadas.
- [ ] Conversa de seguranca (toolbox) concluida e compreendida.
- [ ] EPI e controles de isolamento em vigor.
- [ ] Perigos revistos: ${hazards || "perigos do local/tarefa identificados"}.

4. Verificacoes Durante a Tarefa
- [ ] Trabalho segue o metodo aprovado.
- [ ] Controles mantidos: ${controls || "EPI, isolamento de area e lockout"}.
- [ ] Desvios interrompidos e reportados imediatamente.

5. Encerramento Pos-Tarefa
- [ ] Area de trabalho limpa e segura.
- [ ] Equipamento devolvido a condicao segura normal.
- [ ] Registos de achados e acoes atualizados.

6. Assinatura
- Supervisor: ____________________  Data: __________
- Lider de Equipa: _______________  Data: __________

Notas Adicionais
${extraNotes || "-"}
`.trim();

    const checklistBody = language === "af"
      ? checklistAf
      : language === "zu"
        ? checklistZu
        : language === "pt"
          ? checklistPt
          : checklistEn;

    return `${top}\n\n${checklistBody}`;
  }

  const bodyEn = `
1. Purpose
${scope || "Define the purpose and intended outcome of this document."}

2. Scope
This document applies to all relevant personnel and contractors on site.

3. Responsibilities
- Supervisor: enforce procedure and confirm compliance.
- Team members: follow steps and report deviations.
- HSE/Management: review and approve updates.

4. Hazards / Risks
${hazards || "- Identify hazard exposure before task execution."}

5. Controls
${controls || "- Apply required controls, PPE, and permit steps before start."}

6. Procedure
- Pre-start safety check and communication.
- Execute task using approved method.
- Record outcomes, deviations, and corrective actions.
- Close out and sign off.

7. Records and Evidence
- Keep completed forms/checklists in site records.
- Attach supporting photos where required.

8. Revision Control
Changes must be reviewed and approved before implementation.

Additional Notes
${extraNotes || "-"}
`.trim();

  const bodyAf = `
1. Doel
${scope || "Definieer die doel en gewenste uitkoms van hierdie dokument."}

2. Omvang
Hierdie dokument is van toepassing op relevante personeel en kontrakteurs op die terrein.

3. Verantwoordelikhede
- Toesighouer: pas prosedure toe en bevestig nakoming.
- Spanlede: volg stappe en rapporteer afwykings.
- HSE/Bestuur: hersien en keur opdaterings goed.

4. Gevare / Risiko's
${hazards || "- Identifiseer risiko-blootstelling voor werk begin."}

5. Beheermaatreels
${controls || "- Pas vereiste beheermaatreels, PPE en permit-stappe toe voor aanvang."}

6. Prosedure
- Vooraf veiligheidskontrole en kommunikasie.
- Voer taak uit volgens goedgekeurde metode.
- Teken uitkomste, afwykings en regstellende aksies aan.
- Sluit af en teken af.

7. Rekords en Bewyse
- Bewaar voltooide vorms/kontrolelyste in terreinrekords.
- Heg ondersteunende fotos aan waar nodig.

8. Hersieningsbeheer
Veranderinge moet hersien en goedgekeur word voor implementering.

Bykomende Notas
${extraNotes || "-"}
`.trim();

  const bodyZu = `
1. Inhloso
${scope || "Chaza inhloso nomphumela olindelekile walolu xwebhu."}

2. Ububanzi
Lolu xwebhu lusebenza kubasebenzi nakonkontileka abafanele esizeni.

3. Izibopho
- Umphathi: uqinisekisa ukulandela inqubo.
- Ithimba: lilandela izinyathelo futhi libike ukuphambuka.
- HSE/Abaphathi: babuyekeza futhi bagunyaze izinguquko.

4. Izingozi
${hazards || "- Khomba izingozi ngaphambi kokuqala umsebenzi."}

5. Izilawuli
${controls || "- Sebenzisa izilawuli, i-PPE, nemvume efanele ngaphambi kokuqala."}

6. Inqubo
- Hlola ukuphepha ngaphambi kokuqala.
- Yenza umsebenzi ngendlela evunyelwe.
- Rekhoda imiphumela, ukuphambuka, nezinyathelo zokulungisa.
- Qedela futhi usayine.

7. Amarekhodi Nobufakazi
- Gcina amafomu/ama-checklist aqediwe kumarekhodi esiza.
- Namathisela izithombe lapho kudingeka.

8. Ukulawulwa Kwezibuyekezo
Izinguquko kumele zibuyekezwe futhi zigunyazwe ngaphambi kokusetshenziswa.

Amanothi Engeziwe
${extraNotes || "-"}
`.trim();

  const bodyPt = `
1. Objetivo
${scope || "Definir o objetivo e o resultado esperado deste documento."}

2. Escopo
Este documento aplica-se a todo o pessoal e contratados relevantes no site.

3. Responsabilidades
- Supervisor: aplicar o procedimento e confirmar conformidade.
- Equipa: seguir os passos e reportar desvios.
- HSE/Gestao: rever e aprovar atualizacoes.

4. Perigos / Riscos
${hazards || "- Identificar exposicao ao risco antes da execucao da tarefa."}

5. Controles
${controls || "- Aplicar controles necessarios, EPI e passos de permissao antes do inicio."}

6. Procedimento
- Verificacao de seguranca e comunicacao antes do inicio.
- Executar a tarefa usando o metodo aprovado.
- Registar resultados, desvios e acoes corretivas.
- Encerrar e assinar.

7. Registos e Evidencias
- Guardar formularios/checklists concluidos nos registos do site.
- Anexar fotos de suporte quando necessario.

8. Controlo de Revisao
Alteracoes devem ser revistas e aprovadas antes da implementacao.

Notas Adicionais
${extraNotes || "-"}
`.trim();

  const body = language === "af" ? bodyAf : language === "zu" ? bodyZu : language === "pt" ? bodyPt : bodyEn;
  return `${top}\n\n${body}`;
}

export default async function docsRoutes(app) {
  const dataRoot = process.env.IRONLOG_DATA_DIR || process.cwd();
  db.exec(`
    CREATE TABLE IF NOT EXISTS doc_headers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL,
      site_name TEXT,
      department TEXT,
      prepared_by TEXT,
      approved_by TEXT,
      revision TEXT,
      active INTEGER NOT NULL DEFAULT 1,
      created_by TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT
    );
  `);

  db.exec(`
    CREATE TABLE IF NOT EXISTS doc_drafts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      header_id INTEGER,
      doc_type TEXT NOT NULL,
      title TEXT NOT NULL,
      language TEXT NOT NULL DEFAULT 'en',
      scope_text TEXT,
      hazards_text TEXT,
      controls_text TEXT,
      extra_notes TEXT,
      draft_text TEXT NOT NULL,
      decision TEXT NOT NULL DEFAULT 'pending',
      created_by TEXT,
      decided_by TEXT,
      decided_at TEXT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT,
      FOREIGN KEY(header_id) REFERENCES doc_headers(id)
    );
  `);
  ensureColumn("doc_drafts", "revision_no INTEGER", "revision_no");
  ensureColumn("doc_drafts", "supersedes_draft_id INTEGER", "supersedes_draft_id");

  app.get("/languages", async () => ({
    ok: true,
    rows: [
      { code: "en", name: "English" },
      { code: "af", name: "Afrikaans" },
      { code: "zu", name: "isiZulu" },
      { code: "pt", name: "Portuguese (PT)" },
    ],
  }));

  // Simple diagnostics: which AI provider is configured (without exposing secrets).
  app.get("/ai/status", async () => {
    const useFoundry = toBool(process.env.USE_FOUNDRY || process.env.AI_USE_FOUNDRY || "false");
    const cfg = getAiConfig(useFoundry ? "foundry" : "");
    return {
      ok: true,
      provider: cfg.provider,
      hasAzureEndpoint: Boolean(process.env.AZURE_OPENAI_ENDPOINT),
      hasAzureDeployment: Boolean(process.env.AZURE_OPENAI_DEPLOYMENT),
      hasAzureKey: Boolean(process.env.AZURE_OPENAI_API_KEY),
      hasOpenAIKey: Boolean(process.env.OPENAI_API_KEY),
      hasFoundryKey: Boolean(process.env.FOUNDRY_API_KEY || process.env.AZURE_FOUNDRY_API_KEY || process.env.AZURE_OPENAI_API_KEY),
      hasFoundryEndpoint: Boolean(process.env.FOUNDRY_ENDPOINT || process.env.AZURE_FOUNDRY_ENDPOINT || process.env.AZURE_EXISTING_AIPROJECT_ENDPOINT),
      foundryModel: process.env.FOUNDRY_MODEL || process.env.AZURE_FOUNDRY_MODEL || process.env.OPENAI_MODEL || "gpt-4o-mini",
      model: process.env.OPENAI_MODEL || "gpt-4o-mini",
      azureApiVersion: process.env.AZURE_OPENAI_API_VERSION || "2024-02-15-preview",
    };
  });

  app.post("/ai/review", async (req, reply) => {
    try {
      const body = req.body || {};
      let text = String(body.text || "").trim();
      const filePath = String(body.file_path || "").trim();
      const docType = String(body.doc_type || "General legal/compliance").trim();
      const instructions = String(body.instructions || "Please suggest corrections, retain compliance meaning, and return corrected final text.").trim();
      const preferred = toBool(body.use_foundry);

      if (!text && filePath) {
        const candidatePath = path.isAbsolute(filePath) ? filePath : path.join(dataRoot, filePath);
        if (fs.existsSync(candidatePath) && fs.statSync(candidatePath).isFile()) {
          text = fs.readFileSync(candidatePath, "utf8");
        }
      }

      if (!text) {
        return reply.code(400).send({ ok: false, error: "text or file_path is required" });
      }

      const reviewed = await generateAiDocumentReview({ content: text, docType, instructions, preferFoundry: preferred });
      if (!reviewed) {
        return reply.code(500).send({ ok: false, error: "AI review failed or no output from provider" });
      }

      const outDir = path.join(dataRoot, "uploads", "ai-reviewed");
      fs.mkdirSync(outDir, { recursive: true });

      const outFileName = `${safeFileName(body.filename || docType || "ai-reviewed")}_${Date.now()}.txt`;
      const outFile = path.join(outDir, outFileName);
      fs.writeFileSync(outFile, reviewed, "utf8");

      return reply.send({
        ok: true,
        reviewed_text: reviewed,
        download_url: `/uploads/ai-reviewed/${outFileName}`,
        file_path: outFile,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.post("/ai/ask", async (req, reply) => {
    try {
      const body = req.body || {};
      const machine = String(body.machine || body.asset || "").trim();
      const problem = String(body.problem || body.question || "").trim();
      const context = String(body.context || "").trim();
      const preferred = toBool(body.use_foundry);

      if (!problem) {
        return reply.code(400).send({ ok: false, error: "problem (or question) is required" });
      }

      const answer = await generateAiTechAnswer({
        machine,
        problem,
        context,
        preferFoundry: preferred,
      });

      if (!answer) {
        return reply.code(500).send({ ok: false, error: "AI ask failed or no output from provider" });
      }

      return reply.send({
        ok: true,
        machine: machine || null,
        problem,
        answer,
      });
    } catch (err) {
      req.log.error(err);
      return reply.code(500).send({ ok: false, error: err.message });
    }
  });

  app.get("/headers", async () => {
    const rows = db.prepare(`
      SELECT id, name, site_name, department, prepared_by, approved_by, revision, active, created_by, created_at, updated_at
      FROM doc_headers
      ORDER BY id DESC
      LIMIT 200
    `).all().map((r) => ({ ...r, id: Number(r.id), active: Number(r.active) }));
    return { ok: true, rows };
  });

  app.post("/headers", async (req, reply) => {
    const body = req.body || {};
    const name = String(body.name || "").trim();
    if (!name) return reply.code(400).send({ error: "name is required" });
    const row = {
      name,
      site_name: String(body.site_name || "").trim(),
      department: String(body.department || "").trim(),
      prepared_by: String(body.prepared_by || "").trim(),
      approved_by: String(body.approved_by || "").trim(),
      revision: String(body.revision || "").trim(),
      active: Number(body.active === 0 ? 0 : 1),
      created_by: getUser(req),
      updated_at: nowIso(),
    };
    const res = db.prepare(`
      INSERT INTO doc_headers (name, site_name, department, prepared_by, approved_by, revision, active, created_by, updated_at)
      VALUES (@name, @site_name, @department, @prepared_by, @approved_by, @revision, @active, @created_by, @updated_at)
    `).run(row);
    return { ok: true, id: Number(res.lastInsertRowid) };
  });

  app.get("/drafts", async (req) => {
    const decision = String(req.query?.decision || "").trim().toLowerCase();
    const where = [];
    const params = [];
    if (decision) {
      where.push("d.decision = ?");
      params.push(decision);
    }
    const rows = db.prepare(`
      SELECT
        d.id,
        d.header_id,
        h.name AS header_name,
        d.doc_type,
        d.title,
        d.language,
        d.decision,
        d.revision_no,
        d.supersedes_draft_id,
        d.created_by,
        d.decided_by,
        d.decided_at,
        d.created_at,
        d.updated_at
      FROM doc_drafts d
      LEFT JOIN doc_headers h ON h.id = d.header_id
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY d.id DESC
      LIMIT 300
    `).all(...params).map((r) => ({
      ...r,
      id: Number(r.id),
      header_id: Number(r.header_id || 0) || null,
      revision_no: Number(r.revision_no || 0) || null,
      supersedes_draft_id: Number(r.supersedes_draft_id || 0) || null,
    }));
    return { ok: true, rows };
  });

  app.post("/draft-generate", async (req, reply) => {
    const body = req.body || {};
    const header = resolveHeaderByInputOrLatest(body.header_id);
    if (!header) return reply.code(400).send({ error: "create a header first (none found)" });

    const docType = normalizeDocType(body.doc_type);
    const language = normalizeLang(body.language);
    const useFoundry = toBool(body.use_foundry);
    const title = String(body.title || "").trim();
    if (!title) return reply.code(400).send({ error: "title is required" });

    const scope = String(body.scope || "").trim();
    const hazards = String(body.hazards || "").trim();
    const controls = String(body.controls || "").trim();
    const extraNotes = String(body.extra_notes || "").trim();

    const headerText = buildDraftHeaderText({ docType, title, header });
    const cfg = getAiConfig(useFoundry ? "foundry" : "");
    const aiBody = await generateAIDraftBody({
      language,
      docType,
      scope,
      hazards,
      controls,
      extraNotes,
      preferFoundry: useFoundry,
    });

    const usedAi = Boolean(aiBody);
    const draftText = usedAi
      ? `${headerText}\n\n${aiBody}`
      : draftWithTemplate({
          language,
          docType,
          title,
          header,
          scope,
          hazards,
          controls,
          extraNotes,
        });

    const res = db.prepare(`
      INSERT INTO doc_drafts (
        header_id, doc_type, title, language, scope_text, hazards_text, controls_text, extra_notes,
        draft_text, decision, created_by, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)
    `).run(
      Number(header.id),
      docType,
      title,
      language,
      scope,
      hazards,
      controls,
      extraNotes,
      draftText,
      getUser(req),
      nowIso()
    );
    return {
      ok: true,
      id: Number(res.lastInsertRowid),
      draft_text: draftText,
      ai_provider: cfg.provider || null,
      ai_used: usedAi,
      provider_preference: useFoundry ? "foundry" : "default",
    };
  });

  app.post("/draft-generate-request", async (req, reply) => {
    const body = req.body || {};
    const requestText = String(body.request_text || "").trim();
    if (!requestText) return reply.code(400).send({ error: "request_text is required" });

    const header = resolveHeaderByInputOrLatest(body.header_id);
    if (!header) return reply.code(400).send({ error: "create a header first (none found)" });

    const language = normalizeLang(body.language);
    const docType = normalizeDocType(body.doc_type);
    const title = String(body.title || requestText.slice(0, 80)).trim();
    const useFoundry = toBool(body.use_foundry);
    const useWeb = toBool(body.use_web);
    const trustedOnly = toBool(body.web_trusted_only);
    const trustedDomains = parseTrustedDomains(body.web_trusted_domains || process.env.DOC_WEB_TRUSTED_DOMAINS || "");
    const relatedDocs = findRelevantLegalDocs(requestText, 5);
    const webContext = useWeb ? await fetchWebSearchContext(requestText, 3, trustedOnly ? trustedDomains : []) : [];
    const contextText = relatedDocs.length
      ? relatedDocs
          .map((d) => `- [#${d.id}] ${d.title || "-"} (${d.doc_type || "-"}, ${d.department || "-"})`)
          .join("\n")
      : "- No close match found in current legal library.";
    const webText = webContext.length
      ? webContext.map((w, i) => `- [Web ${i + 1}] ${w.url}\n  ${w.snippet}`).join("\n")
      : "- No web context used.";

    const mergedNotes = `User request:\n${requestText}\n\nRelated legal docs:\n${contextText}\n\nWeb sources:\n${webText}`;
    const authorNotes = String(body.extra_notes || "").trim();
    const aiNotes = [
      authorNotes,
      "The following context is guidance only and should not be copied verbatim into the final document:",
      mergedNotes,
    ].filter(Boolean).join("\n\n");
    const scope = String(body.scope || "").trim();
    const hazards = String(body.hazards || "").trim();
    const controls = String(body.controls || "").trim();

    const headerText = buildDraftHeaderText({ docType, title, header });
    const cfg = getAiConfig();
    const aiBody = await generateAIDraftBody({
      language,
      docType,
      scope: scope || "",
      hazards,
      controls,
      extraNotes: aiNotes,
      preferFoundry: useFoundry,
    });

    const usedAi = Boolean(aiBody);
    const draftText = usedAi
      ? `${headerText}\n\n${aiBody}`
      : draftWithTemplate({
          language,
          docType,
          title,
          header,
          scope: scope || "",
          hazards,
          controls,
          extraNotes: authorNotes,
        });

    const res = db.prepare(`
      INSERT INTO doc_drafts (
        header_id, doc_type, title, language, scope_text, hazards_text, controls_text, extra_notes,
        draft_text, decision, created_by, updated_at
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)
    `).run(
      Number(header.id),
      docType,
      title,
      language,
      scope || "",
      hazards,
      controls,
      authorNotes,
      draftText,
      getUser(req),
      nowIso()
    );

    return {
      ok: true,
      id: Number(res.lastInsertRowid),
      draft_text: draftText,
      ai_provider: cfg.provider || null,
      ai_used: usedAi,
      provider_preference: useFoundry ? "foundry" : "default",
      web_used: Boolean(webContext.length),
      web_trusted_only: trustedOnly,
      web_trusted_domains: trustedOnly ? trustedDomains : [],
      related_docs: relatedDocs.map((d) => ({
        id: Number(d.id),
        title: d.title || "",
        doc_type: d.doc_type || "",
        department: d.department || "",
      })),
      web_sources: webContext.map((w) => ({ url: w.url })),
    };
  });

  app.post("/drafts/:id/decision", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ error: "invalid id" });
    const approved = Boolean(req.body?.approved);
    const next = approved ? "approved" : "rejected";
    const user = getUser(req);
    const ts = nowIso();

    const current = db.prepare(`
      SELECT id, header_id, doc_type, title, language, decision
      FROM doc_drafts
      WHERE id = ?
    `).get(id);
    if (!current) return reply.code(404).send({ error: "draft not found" });

    if (!approved) {
      const updated = db.prepare(`
        UPDATE doc_drafts
        SET decision = ?, decided_by = ?, decided_at = ?, updated_at = ?
        WHERE id = ?
      `).run(next, user, ts, ts, id);
      if (!updated.changes) return reply.code(404).send({ error: "draft not found" });
      return { ok: true, id, decision: next };
    }

    // Auto-versioning + supersede flow:
    // - Find latest currently approved doc in same family
    // - New approved doc gets revision_no = prior + 1
    // - Prior approved doc becomes superseded
    const prior = db.prepare(`
      SELECT id, revision_no
      FROM doc_drafts
      WHERE id <> ?
        AND header_id IS ?
        AND doc_type = ?
        AND title = ?
        AND language = ?
        AND decision = 'approved'
      ORDER BY COALESCE(revision_no, 0) DESC, id DESC
      LIMIT 1
    `).get(
      id,
      current.header_id ?? null,
      String(current.doc_type || ""),
      String(current.title || ""),
      String(current.language || "en")
    );

    const nextRevision = Number(prior?.revision_no || 0) + 1;
    const tx = db.transaction(() => {
      if (prior?.id) {
        db.prepare(`
          UPDATE doc_drafts
          SET decision = 'superseded', updated_at = ?
          WHERE id = ?
        `).run(ts, Number(prior.id));
      }

      const updated = db.prepare(`
        UPDATE doc_drafts
        SET
          decision = 'approved',
          revision_no = ?,
          supersedes_draft_id = ?,
          decided_by = ?,
          decided_at = ?,
          updated_at = ?
        WHERE id = ?
      `).run(nextRevision, prior?.id ? Number(prior.id) : null, user, ts, ts, id);
      if (!updated.changes) throw new Error("draft not found");
    });
    try {
      tx();
      return {
        ok: true,
        id,
        decision: "approved",
        revision_no: nextRevision,
        supersedes_draft_id: prior?.id ? Number(prior.id) : null,
      };
    } catch (e) {
      return reply.code(500).send({ error: e?.message || "approval update failed" });
    }
  });

  app.get("/drafts/:id.pdf", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ error: "invalid id" });

    const row = db.prepare(`
      SELECT
        d.id, d.header_id, d.doc_type, d.title, d.language, d.scope_text, d.hazards_text, d.controls_text,
        d.extra_notes, d.draft_text, d.decision, d.revision_no, d.created_by, d.decided_by, d.decided_at, d.created_at,
        h.name AS header_name, h.site_name, h.department, h.prepared_by, h.approved_by, h.revision
      FROM doc_drafts d
      LEFT JOIN doc_headers h ON h.id = d.header_id
      WHERE d.id = ?
    `).get(id);
    if (!row) return reply.code(404).send({ error: "draft not found" });
    if (String(row.decision || "").toLowerCase() !== "approved") {
      return reply.code(409).send({
        error: "only approved documents can be exported to PDF",
        decision: row.decision || "pending",
      });
    }

    const pdf = await buildPdfBuffer((doc) => {
      sectionTitle(doc, `${row.doc_type || "Document"} - ${row.title || ""}`);

      kvGrid(doc, [
        { k: "Draft ID", v: String(row.id) },
        { k: "Revision", v: row.revision_no ? `Rev ${row.revision_no}` : "-" },
        { k: "Decision", v: String(row.decision || "pending").toUpperCase() },
        { k: "Language", v: String(row.language || "en").toUpperCase() },
        { k: "Header", v: row.header_name || "-" },
        { k: "Site", v: row.site_name || "-" },
        { k: "Department", v: row.department || "-" },
        { k: "Prepared By", v: row.prepared_by || "-" },
        { k: "Approved By", v: row.approved_by || "-" },
        { k: "Revision", v: row.revision || "-" },
        { k: "Created By", v: row.created_by || "-" },
        { k: "Created At", v: row.created_at || "-" },
      ]);

      ensurePageSpace(doc, 40);
      sectionTitle(doc, "Document Body");
      const bodyOnly = stripHeaderFromDraftText(row.draft_text);
      doc.font("Helvetica").fontSize(10).fillColor("#111111");
      doc.text(bodyOnly, {
        width: doc.page.width - doc.page.margins.left - doc.page.margins.right,
        align: "left",
      });
    }, {
      title: "IRONLOG",
      subtitle: "AI Document - Final Issue",
      rightText: `Draft #${row.id}`,
    });

    const safeTitle = safeFileName(row.title, `draft-${id}`);
    const download = String(req.query?.download || "0") === "1";
    reply
      .header("Content-Type", "application/pdf")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename=\"${safeTitle}.pdf\"`)
      .send(pdf);
  });

  app.get("/drafts/:id.docx", async (req, reply) => {
    const id = Number(req.params?.id || 0);
    if (!id) return reply.code(400).send({ error: "invalid id" });

    const row = db.prepare(`
      SELECT
        d.id, d.header_id, d.doc_type, d.title, d.language, d.scope_text, d.hazards_text, d.controls_text,
        d.extra_notes, d.draft_text, d.decision, d.revision_no, d.created_by, d.decided_by, d.decided_at, d.created_at,
        h.name AS header_name, h.site_name, h.department, h.prepared_by, h.approved_by, h.revision
      FROM doc_drafts d
      LEFT JOIN doc_headers h ON h.id = d.header_id
      WHERE d.id = ?
    `).get(id);
    if (!row) return reply.code(404).send({ error: "draft not found" });
    if (String(row.decision || "").toLowerCase() !== "approved") {
      return reply.code(409).send({
        error: "only approved documents can be exported to Word",
        decision: row.decision || "pending",
      });
    }

    const bodyLines = splitLines(stripHeaderFromDraftText(row.draft_text));
    const doc = new Document({
      sections: [{
        children: [
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun(`${row.doc_type || "Document"} - ${row.title || ""}`)],
          }),
          new Paragraph({ children: [new TextRun(`Draft ID: ${row.id}`)] }),
          new Paragraph({ children: [new TextRun(`Revision: ${row.revision_no ? `Rev ${row.revision_no}` : "-"}`)] }),
          new Paragraph({ children: [new TextRun(`Decision: ${String(row.decision || "pending").toUpperCase()}`)] }),
          new Paragraph({ children: [new TextRun(`Language: ${String(row.language || "en").toUpperCase()}`)] }),
          new Paragraph({ children: [new TextRun(`Header: ${row.header_name || "-"}`)] }),
          new Paragraph({ children: [new TextRun(`Site: ${row.site_name || "-"}`)] }),
          new Paragraph({ children: [new TextRun(`Department: ${row.department || "-"}`)] }),
          new Paragraph({ children: [new TextRun(`Prepared By: ${row.prepared_by || "-"}`)] }),
          new Paragraph({ children: [new TextRun(`Approved By: ${row.approved_by || "-"}`)] }),
          new Paragraph({ children: [new TextRun(`Created By: ${row.created_by || "-"}`)] }),
          new Paragraph({ children: [new TextRun(`Created At: ${row.created_at || "-"}`)] }),
          new Paragraph({ children: [new TextRun("")] }),
          new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Document Body")] }),
          ...bodyLines.map((line) => new Paragraph({ children: [new TextRun(line || " ")] })),
        ],
      }],
    });
    const buffer = await Packer.toBuffer(doc);
    const safeTitle = safeFileName(row.title, `draft-${id}`);
    const download = toBool(req.query?.download);
    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
      .header("Content-Disposition", `${download ? "attachment" : "inline"}; filename=\"${safeTitle}.docx\"`)
      .send(buffer);
  });

  app.get("/register.pdf", async (req, reply) => {
    const currentOnly = toBool(req.query?.current_only);
    const rows = rowsForApprovedRegister(currentOnly);

    const pdf = await buildPdfBuffer((doc) => {
      sectionTitle(doc, currentOnly ? "Approved Documents Register (Current Only)" : "Approved Documents Register");
      kvGrid(doc, [
        { k: "Approved Documents", v: String(rows.length) },
        { k: "Current Only", v: currentOnly ? "Yes" : "No" },
        { k: "Generated At", v: nowIso() },
      ]);

      ensurePageSpace(doc, 30);
      sectionTitle(doc, "Documents");
      const left = doc.page.margins.left;
      const width = doc.page.width - doc.page.margins.left - doc.page.margins.right;

      if (!rows.length) {
        doc.font("Helvetica").fontSize(10).text("No approved documents found.");
        return;
      }

      // Simple manual rows to keep this endpoint lightweight.
      rows.forEach((r) => {
        ensurePageSpace(doc, 36);
        const rev = r.revision_no ? `Rev ${r.revision_no}` : "-";
        const line1 = `#${r.id} | ${r.doc_type || "-"} | ${r.title || "-"}`;
        const line2 = `${rev} | ${String(r.language || "en").toUpperCase()} | ${r.department || "-"} | ${r.site_name || "-"} | Approved by ${r.decided_by || "-"} on ${r.decided_at || "-"}`;
        doc.font("Helvetica-Bold").fontSize(9).text(line1, left, doc.y, { width });
        doc.font("Helvetica").fontSize(9).text(line2, left, doc.y + 1, { width });
        doc.moveDown(0.4);
      });
    }, {
      title: "IRONLOG",
      subtitle: "Approved Documents Register",
      rightText: "Compliance Export",
    });

    const download = String(req.query?.download || "0") === "1";
    reply
      .header("Content-Type", "application/pdf")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename=\"approved_documents_register.pdf\"`
      )
      .send(pdf);
  });

  app.get("/register.docx", async (req, reply) => {
    const currentOnly = toBool(req.query?.current_only);
    const rows = rowsForApprovedRegister(currentOnly);
    const children = [
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun(currentOnly ? "Approved Documents Register (Current Only)" : "Approved Documents Register")],
      }),
      new Paragraph({ children: [new TextRun(`Approved Documents: ${rows.length}`)] }),
      new Paragraph({ children: [new TextRun(`Current Only: ${currentOnly ? "Yes" : "No"}`)] }),
      new Paragraph({ children: [new TextRun(`Generated At: ${nowIso()}`)] }),
      new Paragraph({ children: [new TextRun("")] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Documents")] }),
    ];
    if (!rows.length) {
      children.push(new Paragraph({ children: [new TextRun("No approved documents found.")] }));
    } else {
      rows.forEach((r) => {
        const rev = r.revision_no ? `Rev ${r.revision_no}` : "-";
        children.push(new Paragraph({
          children: [new TextRun(`#${r.id} | ${r.doc_type || "-"} | ${r.title || "-"}`)],
        }));
        children.push(new Paragraph({
          children: [new TextRun(`${rev} | ${String(r.language || "en").toUpperCase()} | ${r.department || "-"} | ${r.site_name || "-"} | Approved by ${r.decided_by || "-"} on ${r.decided_at || "-"}`)],
        }));
      });
    }

    const doc = new Document({ sections: [{ children }] });
    const buffer = await Packer.toBuffer(doc);
    const download = toBool(req.query?.download);
    return reply
      .header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
      .header(
        "Content-Disposition",
        `${download ? "attachment" : "inline"}; filename=\"approved_documents_register.docx\"`
      )
      .send(buffer);
  });
}

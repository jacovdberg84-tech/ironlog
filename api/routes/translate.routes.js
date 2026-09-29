// IRONLOG/api/routes/translate.routes.js — POST /api/translate { text, to: "en" | "pt" }
import { aiChatText, isAiConfigured } from "../utils/ironmind.js";
import { translateText } from "../utils/translate.js";

export default async function translateRoutes(app) {
  app.post("/", async (req, reply) => {
    const text = String(req.body?.text || "").slice(0, 4000);
    const to = String(req.body?.to || "en").toLowerCase() === "pt" ? "pt" : "en";
    if (!text.trim()) return reply.code(400).send({ ok: false, error: "text is required" });
    if (!isAiConfigured()) return { ok: false, available: false, error: "AI translation is not set up on this server" };
    const out = await translateText(text, { to, chat: aiChatText });
    if (out == null) return reply.code(502).send({ ok: false, available: true, error: "Translation failed — try again" });
    return { ok: true, available: true, to, text: out };
  });
}

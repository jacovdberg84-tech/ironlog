// IRONLOG/api/routes/assets/qr-profiles.routes.js — Asset, undercarriage and tyre QR profiles.
// Registered by routes/assets.routes.js; shared helpers arrive through ctx.

export default function registerQrProfilesRoutes(app, ctx) {
  const {
    buildQrProfile,
    buildTyreQrProfile,
    buildUndercarriageQrProfile,
    getAssetByCode,
    getStoredQrProfile,
    getStoredTyreQrProfile,
    getStoredUndercarriageQrProfile,
    upsertQrProfile,
    upsertTyreQrProfile,
    upsertUndercarriageQrProfile,
  } = ctx;

  app.get("/:asset_code/qr-profile", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "Asset code is required" });
    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const stored = getStoredQrProfile.get(asset.id);
    const live = buildQrProfile(asset, req);
    let storedPayload = null;
    if (stored?.qr_payload) {
      try {
        storedPayload = JSON.parse(String(stored.qr_payload || "{}"));
      } catch {
        storedPayload = null;
      }
    }

    return {
      ok: true,
      asset_code: asset.asset_code,
      stored: stored
        ? {
            qr_payload: storedPayload,
            qr_text: stored.qr_text,
            generated_at: stored.generated_at,
          }
        : null,
      live_preview: live.profile,
      live_qr_text: live.qrText,
    };
  });

  app.post("/:asset_code/qr-profile/refresh", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "Asset code is required" });
    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const built = buildQrProfile(asset, req);
    upsertQrProfile.run(asset.id, JSON.stringify(built.profile), built.qrText);

    return {
      ok: true,
      asset_code: asset.asset_code,
      qr_payload: built.profile,
      qr_text: built.qrText,
    };
  });

  app.get("/:asset_code/undercarriage-qr-profile", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "Asset code is required" });
    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const stored = getStoredUndercarriageQrProfile.get(asset.id);
    let storedPayload = null;
    if (stored?.qr_payload) {
      try {
        storedPayload = JSON.parse(String(stored.qr_payload || "{}"));
      } catch {
        storedPayload = null;
      }
    }
    const live = buildUndercarriageQrProfile(asset, req);
    return {
      ok: true,
      asset_code: asset.asset_code,
      stored: storedPayload
        ? { qr_payload: storedPayload, qr_text: stored.qr_text, generated_at: stored.generated_at }
        : null,
      live_preview: live.profile,
      live_qr_text: live.qrText,
    };
  });

  app.post("/:asset_code/undercarriage-qr-profile/refresh", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "Asset code is required" });
    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const built = buildUndercarriageQrProfile(asset, req);
    upsertUndercarriageQrProfile.run(asset.id, JSON.stringify(built.profile), built.qrText);

    return {
      ok: true,
      asset_code: asset.asset_code,
      qr_payload: built.profile,
      qr_text: built.qrText,
    };
  });

  app.get("/:asset_code/tyre-qr-profile", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "Asset code is required" });
    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const stored = getStoredTyreQrProfile.get(asset.id);
    let storedPayload = null;
    if (stored?.qr_payload) {
      try {
        storedPayload = JSON.parse(String(stored.qr_payload || "{}"));
      } catch {
        storedPayload = null;
      }
    }
    const live = buildTyreQrProfile(asset, req);
    return {
      ok: true,
      asset_code: asset.asset_code,
      stored: storedPayload
        ? { qr_payload: storedPayload, qr_text: stored.qr_text, generated_at: stored.generated_at }
        : null,
      live_preview: live.profile,
      live_qr_text: live.qrText,
    };
  });

  app.post("/:asset_code/tyre-qr-profile/refresh", async (req, reply) => {
    const asset_code = String(req.params.asset_code || "").trim();
    if (!asset_code) return reply.code(400).send({ error: "Asset code is required" });
    const asset = getAssetByCode.get(asset_code);
    if (!asset) return reply.code(404).send({ error: "Asset not found" });

    const built = buildTyreQrProfile(asset, req);
    upsertTyreQrProfile.run(asset.id, JSON.stringify(built.profile), built.qrText);

    return {
      ok: true,
      asset_code: asset.asset_code,
      qr_payload: built.profile,
      qr_text: built.qrText,
    };
  });
}

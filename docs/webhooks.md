# Webhooks

IronLog sends a signed JSON message to another system when something happens (`api/utils/webhooks.js`, routes in `api/routes/webhooks.routes.js`, screen under Admin → Webhooks, receiver guide on `web/api-docs.html#webhooks`).

- **Events:** `breakdown.created`, `breakdown.closed`, `workorder.created`, `workorder.completed`, `stock.below_minimum`. A webhook can be limited to sites; stock is company-wide and goes to every subscribed webhook.
- **How events are found:** a scan every 30 seconds checks the tables. A breakdown from any screen, the technician portal, the API or an import is caught the same way, and no screen code changes. `webhook_seen` remembers what has been seen. Records from before a subscription are never sent. A part warns again only after it has been restocked.
- **Signing:** `X-IronLog-Signature: t=<unix>,v1=<HMAC-SHA256("<t>.<body>", secret)>`. The secret (`whsec_…`) is shown once, stored encrypted, and can be replaced ("New secret").
- **Delivery:** answer 2xx within 10 seconds. Otherwise IronLog retries after 1 min, 5 min, 30 min, 2 h, 6 h and 12 h, then marks the message failed. Admins see each message and can send it again. Redirects are not followed. Delivered messages are kept for 60 days.
- **Addresses:** https only, and private or local addresses are refused, both when saving and when sending (checked through DNS). For an on-site system, set `IRONLOG_WEBHOOK_ALLOW_PRIVATE=1` on the server; this also allows http.
- **Who:** administrators only. Every change is audited. Webhooks are part of the licensed `api_access` module (Enterprise), together with API keys.

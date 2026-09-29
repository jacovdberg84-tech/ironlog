# Borris: maintenance planning and costing

## Costing gaps (My Work)
Plain database checks, no AI needed (`api/utils/costingGaps.js`):

| Gap | Meaning |
| --- | --- |
| `service_unpriced` | Service due within 250 h/km (or overdue) with no price from a manual input, template or past services |
| `part_zero_cost` | Store part at $0 that is in an active service kit or was issued in the last 90 days |
| `service_no_labour` | Service work order finished in the last 60 days with no labour hours |
| `labour_rate_missing` | No default labour rate in cost settings (IRONLOG falls back to $35/h) |

"Not needed" stores the gap in `costing_gap_dismissals` with a reason and user.

## Borris costing assistant
`POST /api/maintenance/costing-gaps/assist { key }`

1. IRONLOG gathers evidence: this machine's service history, sister machines on
   the same model (or category) and interval, store items that mention the model,
   oils, matching Workshop Library pages and the standard labour hours.
2. Borris picks parts and quantities from that evidence and replies in JSON.
3. IRONLOG drops any part code that is not in stores, prices every line from the
   store, and falls back to service history when Borris is offline or unsure.
4. The person edits quantities or labour and clicks **Apply**. Services are saved
   as weekly-forum service cost inputs; part prices through the part-cost endpoint.

Nothing is written without that click.

## Ollama settings
Borris runs on the server's Ollama. Costing proposals and planning questions use
Ollama's `/api/chat` in JSON mode.

- `BORRIS_NUM_CTX` (default off): a larger context window for Borris, e.g. `8192`.
  Ollama reloads the model with more memory when the window changes, so only set
  it when the server has the RAM to spare (roughly +1–2 GB for a 7–8B model at
  8192). On 2026-09-29 the default of 8192 exhausted the server and dropped the
  Cloudflare tunnel, which is why it is now off.
- `BORRIS_COSTING_TIMEOUT_MS` (default `45000`, capped at `80000`): how long a
  costing proposal may take before IRONLOG falls back to the service-history
  proposal. Cloudflare cuts requests at about 100 s.
- One costing proposal runs at a time; a second request gets the history
  proposal straight away.
- A model that follows JSON instructions well (e.g. `qwen2.5:7b` or `llama3.1:8b`)
  gives better proposals than very small models.

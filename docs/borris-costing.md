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
Borris runs on the server's Ollama. Planning questions and costing proposals use
Ollama's `/api/chat` with a larger context window and JSON mode.

- `BORRIS_NUM_CTX` (default `8192`): tokens Borris can read at once. `16384` gives
  more room for evidence but needs roughly 1–2 GB more RAM for a 7–8B model.
  `0` keeps the model's own default.
- `BORRIS_COSTING_TIMEOUT_MS` (default `120000`): how long a costing proposal may take.
- A model that follows JSON instructions well (e.g. `qwen2.5:7b` or `llama3.1:8b`)
  gives better proposals than very small models.

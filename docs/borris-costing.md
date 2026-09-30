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
`POST /api/maintenance/costing-gaps/assist { key }` answers straight away with a
proposal from the records (the machine's history, else sister machines), and
queues Borris in the background. The panel polls
`GET /costing-gaps/assist/status?key=` and offers **Use Borris's proposal** when
he is done.

- Evidence: this machine's service history, sister machines on the same model (or
  category) and interval, store items that mention the model, oils, Workshop
  Library pages whose title/model/applicability names the model, the standard
  labour hours, and the planner's notes.
- **Tell Borris what you know**: notes saved per service plan
  (`costing_planner_notes`, `POST /costing-gaps/notes`). Borris treats them as
  facts first; store part codes typed in the notes join his candidates. Saving
  notes asks Borris again.
- **Add part**: the planner can add any store item to the proposal themselves.
- Borris's reply is checked: part codes not in stores are dropped, every line is
  priced from the store, implausible labour falls back to the standard.
- One Borris job runs at a time (`api/utils/borrisQueue.js`); results are kept
  for 30 minutes. Each opened gap is numbered in the browser so a late answer for
  one machine never shows in another machine's panel.
- Nothing is saved until the person clicks **Apply**.

## Ollama settings
Borris runs on the server's Ollama. Costing proposals and planning questions use
Ollama's `/api/chat` in JSON mode.

- `BORRIS_NUM_CTX` (default off): a larger context window for Borris, e.g. `8192`.
  Ollama reloads the model with more memory when the window changes, so only set
  it when the server has the RAM to spare (roughly +1–2 GB for a 7–8B model at
  8192). On 2026-09-29 the default of 8192 exhausted the server and dropped the
  Cloudflare tunnel, which is why it is now off.
- `BORRIS_COSTING_TIMEOUT_MS` (default `300000` = 5 min, max 15 min): how long
  Borris may think in the background. The web request never waits for him.
- A model that follows JSON instructions well (e.g. `qwen2.5:7b` or `llama3.1:8b`)
  gives better proposals than very small models.

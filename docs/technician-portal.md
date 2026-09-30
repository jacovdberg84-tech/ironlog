# Technician portal — discovery and plan

Status: built (phases 1–7). Sections 1–6 are the discovery and plan; section 7 onward describes what shipped.

## 1. What exists today

| Area | What is there | Reuse |
| --- | --- | --- |
| Portal | `web/technician-terminal.html/.js`: PIN or password login, list of my WOs, "open WO #". Links hidden in Admin text and the Work Orders roster section. Opens `workorder-qr.html` (status buttons). | Keep the URL and login (`auth-shared.js`), rebuild the workspace. |
| Auth | Session token; hook sets `x-user-name / x-user-role / x-user-roles`. `artisan` role. PIN login + roster (`/api/auth/pin-roster`). | As is. |
| Permissions | Roles + named permissions (`ROLE_PERMISSION_FALLBACK`, `requirePermission`); artisan has `workorders.close.request`. `canRoleTransition`: artisan may assigned→in_progress→completed on WOs assigned to them (`technicianMatchesUser`). | Extend the artisan list with `tech.*` keys; keep existing checks. |
| Work orders | `work_orders` (one `assigned_artisan_name`, `started_at`, `completed_at`, `completion_notes`, `repair_progress`, `labor_hours`, `labor_rate_per_hour`, `job_description`, `due_date`, `priority`). Status route `POST /api/workorders/:id/status` enforces transitions, artisan ownership, completion notes, breakdown sync. | The portal changes WO status only through this route. |
| Services | `maintenance_plans`, due list, service templates, `work_order_planned_materials`, `stock_reservations`. | Read-only in the portal. |
| Breakdowns | `breakdowns` (+ component, ETS, parts status), `POST /api/breakdowns`, `PATCH /:id/details`. | Report breakdown from the portal uses the same route. |
| Stock | `parts`, `stock_movements` (qty, location/bin), `stock_bins`, `/api/stock/onhand`, issue to WO = `POST /api/workorders/:id/issue` (stores/supervisor only). | Source of truth; technicians request, stores issue. |
| Parts requests | `maintenance_parts_requests` + `POST /api/maintenance/parts-requests`; stores queue on My Work / Parts Orders. | Request part = this route. |
| Labour | (a) `work_orders.labor_hours` → costing, finance, journals, reports. (b) `mechanic_labor_entries` (mechanics timesheet) → timesheet, monthly costs, and the cost dashboard, which **replaces** an asset's WO labour cost with timesheet cost when present. | See §5 risk. |
| Photos | LDV/machine check photos; no work-order photos. | New WO photo table using the same upload helper. |
| History | `GET /api/assets/:code/history`, asset QR hub (`asset-qr.html`, `/api/assets/:code/qr-profile`). | Asset view reuses both. |
| Manuals | Workshop Library FTS (`searchWorkshop`). | WO/asset "manuals" list. |
| Borris | `/api/ironmind/ask` (advisory), costing assistant. No write tools. | Contextual "Ask Borris" = ask with WO context. |
| Offline | Per-page localStorage queues (pre-starts, safety) with sync banners; `sw.js` exists but is not registered and would cache API responses. | Same queue pattern + idempotency keys; a portal-scoped service worker that caches only the portal shell. |

## 2. Gaps
- No "today" view, no active-job timer, no pause/waiting states (WO has only open/assigned/in_progress/completed).
- Labour is typed once at completion; nothing derives it from work.
- One technician per WO.
- No findings, WO photos, shift report or handover.
- Technicians cannot see stock availability or bins on the WO.
- No asset operational view for technicians; no portal offline capture.

## 3. Architecture
- **UI**: `technician-terminal.html` becomes the portal (same URL, also linked from the menu for artisans). Screens: Today, Work order, Shift, My week, Asset. Plain classic scripts like the rest of `web/`.
- **API**: `/api/tech/*` is a thin read/compose layer over existing tables, plus capture of what does not exist yet (activity events, findings, photos, shifts). Anything that already has a route (WO status, parts request, breakdown, issue) is called through that route's logic, not reimplemented.
- **One source of truth per fact**: WO status stays in `work_orders`; time is derived from activity events; parts from stock; requests from `maintenance_parts_requests`.

### Activity and time
Append-only `tech_activity_events` (wo, user, action, at, note, `client_event_id` unique). Actions: start, pause, resume, waiting_parts, waiting_ops, testing, complete. Time segments are computed from the events, never stored twice:
- `active` and `testing` count as labour; `paused`, `waiting_parts`, `waiting_ops` do not.
- One running job per technician: starting another job pauses the current one.
- WO status mapping: start → `in_progress` (via the status route); complete → `completed` with the technician's notes. Waiting/testing states stay `in_progress` on the WO and show on the board.
- On complete, `labor_hours` is filled from the derived active time only if nobody entered hours (manual entry wins).

### Shift report
`tech_shifts` (user, started_at, ended_at, status draft/submitted, findings, unplanned work, safety, outstanding, handover, next shift). The timeline is built from activity events between start and end, so a night shift crossing midnight works. Submit locks the report and lists it for supervisors.

## 4. Schema (all additive, created with `CREATE TABLE IF NOT EXISTS`)
`tech_activity_events`, `tech_findings`, `work_order_photos`, `tech_shifts`, `work_order_technicians` (helpers). No existing column changes. Rollback: redeploy the previous version; the new tables are simply unused (drop them only if wanted).

## 5. Risks and decisions
- **Labour numbers** (decided 2026-09-30: *job + timesheet*): `labor_hours` is filled from the timer when nobody typed hours, and shift submit writes the shift's job time to `mechanic_labor_entries` (tagged with the WO number). Costing therefore rises to the real figure for jobs that had 0 labour, and the cost dashboard uses the timesheet cost for those machines, as it already does for manual timesheet rows.
- **Helpers** (decided: *lead + helpers*): the assigned technician leads and changes status; helpers added by the foreman log their own time. WO labour = everyone's active time.
- **Shift** (decided: *tap start / submit*): starts at Start shift or the first job, ends at submit; crossing midnight is fine.
- Status rules are unchanged: the portal cannot do anything the status route refuses.
- Stock: no issuing from the portal; technicians request, stores issue.
- Borris stays read-only.

## 6. Phases
1. Today dashboard + assigned work (+ menu link).
2. Work order view + activity/time capture + findings/photos.
3. Parts on the WO: planned, availability, bin, request.
4. Shift report + handover (+ supervisor list).
5. My week + QR asset view.
6. Offline queue + portal service worker.
7. Contextual Borris.

## 7. What shipped

### Files
- API: `api/routes/tech.routes.js` (`/api/tech`), `api/utils/techActivity.js` (states, time segments, schema), `api/utils/techPortal.js` (read-only composition of work orders, stock, requests), `api/utils/techShift.js` (shift timeline, timesheet rows), `api/utils/technicianIdentity.js` (shared "is this my job" match, moved out of `workorders.routes.js` unchanged). Registered in `api/server.js`.
- Web: `web/technician-terminal.html/.js` (same URL and PIN login), `web/tech-portal.js/.css`, `web/tech-sw.js` (portal-only service worker), `web/tech-manifest.json`. Menu link "Technician portal" for workshop roles (`index.html`, `app/core.js`); "Technician view" button on the machine QR page for signed-in workshop staff (`asset-qr.html`).
- Tests: `api/test/techPortal.test.js`.

### Routes (`/api/tech`)
| Route | What it does |
| --- | --- |
| `GET /today` | My jobs grouped urgent / planned / waiting / completed today, current running job, shift, notifications (new job, part arrived). |
| `GET /workorders/:id` | Job view: WO, machine + meter, breakdown or service, team and states, my time, parts, findings, photos, earlier jobs, activity. |
| `POST /workorders/:id/action` | start, pause, resume, waiting_parts, waiting_ops, testing, complete. Start/complete change the WO through `POST /api/workorders/:id/status`. |
| `POST /findings` | Finding / safety / note on a job or a machine. |
| `POST /workorders/:id/photos` | Job photo (multipart). |
| `POST /workorders/:id/parts-request` | Calls `POST /api/maintenance/parts-requests`. |
| `GET /parts/search` | Stores items with on-hand and bin. |
| `POST/DELETE /workorders/:id/helpers` | Foreman adds/removes helpers. |
| `GET/PUT /shift`, `POST /shift/start`, `POST /shift/submit`, `GET /shifts` | Shift report, handover, submit, team list. |
| `GET /week` | Hours per day, jobs completed/open/waiting. |
| `GET /assets/:code`, `POST /breakdowns` | Machine view; breakdown report via `POST /api/breakdowns`. |

### Permissions
- Portal: artisan, supervisor, workshop_admin, admin, plant_manager, site_manager. Operators and stores get 403.
- A technician sees and acts only on jobs where they are the lead (assigned) or a helper. Reassigning a job removes the old technician's access at once.
- Only the lead (or a foreman) starts an assigned job and completes the WO; helpers finish their own part. Adding helpers: foreman roles only.
- Unchanged: technicians request parts (they cannot issue stock), cannot close/approve, delete, change pricing, service intervals, costing or users. Every WO status change passes the existing status route and its rules.
- Borris on the job screen uses the existing read-only `/api/ironmind/ask` with the job as context; it cannot write anything.

### Labour
- `work_orders.labor_hours` on completion = hours typed in the complete sheet, else the timer total (everyone's working + testing time) — only when the WO had no hours already.
- Shift submit writes one `mechanic_labor_entries` row per job worked (hours, start/finish clock, category, `job_card_no` = WO number, `created_by` = `portal:<user>`). Waiting and paused time is never labour.
- Impact: jobs that used to close with 0 labour now carry their real hours, and the cost dashboard uses timesheet cost for those machines (as it already does for manual timesheet rows). Existing records are not changed.

### Offline
- Writes (actions, findings, photos, part requests, breakdowns, shift start/submit) carry a `client_event_id` and the time they happened. Without signal they are kept in the phone's storage (`ironlog-tech-queue-v1`) and sent in order when the signal returns (also every minute, and on "Send now"). The server stores each `client_event_id` once (`tech_client_events`), so a retry never saves twice. Offline times older than 72 h or in the future fall back to server time.
- Last-seen screens are kept on the phone and shown with a "no signal" note. A refusal (for example the job was reassigned) is shown in the sync bar and not retried forever.
- `tech-sw.js` caches only the portal's own page, scripts and styles (network first) and never API data. The old unregistered `web/sw.js` is untouched.

### Tests
- Unit: state rules, labour excludes waiting, open jobs, night shift over midnight → one report and one timesheet row per job, offline timestamp window.
- API flow: role gate, Today, start (WO → in progress), duplicate offline event saved once, helper added by foreman only, finding, stores search, part request (no stock issued), waiting / resume, completion needs notes, labour 4.5 h (lead + helper), breakdown closed, shift report with handover, timesheet row, duplicate submit, team list, week, machine view.
- Edge cases: WO reassignment, failed stores request then retry with the same offline id, helper finishing their part leaves the WO open, abandoned running job stopped at shift submit.

### Limitations
- Helpers are managed from the portal job screen (foreman), not yet from the Work Orders board drawer.
- Timesheet rows are written at shift submit; a shift never submitted writes no timesheet rows (WO labour is still filled on completion).
- Photos taken offline are kept in phone storage (about 10 photos before it is full).
- The meter-hours lookup is still duplicated in three older route files (unchanged here).

### Migration and rollback
- New tables only, created with `CREATE TABLE IF NOT EXISTS` at start-up: `tech_activity_events`, `tech_findings`, `work_order_photos`, `work_order_technicians`, `tech_shifts`, `tech_client_events`. The timesheet's optional columns (`category`, `time_started`, `time_finished`, `job_card_no`) are added if missing, as the timesheet screen already does.
- Rollback: redeploy the previous release. The new tables are then unused; drop them only if wanted. Timesheet rows written by the portal are marked `created_by = 'portal:<user>'` and can be removed with that filter.

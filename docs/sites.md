# Sites

Every record belongs to a site (`site_code`; `main` when a company has one site,
or `IRONLOG_DEFAULT_SITE_CODE`). Each IronLog company (AML, the pilot…) has its
own database, so companies never share records; sites split one company's work.

- **Who sees what.** `users.allowed_locations` lists the sites a person may use;
  empty means all sites. The sign-in check (`api/auth/hook.js`) refuses any other
  site and sets `x-site-code`; administrators may use every site. Routes read the
  chosen site with `getSiteCode(req)` (`api/utils/request.js`).
- **Real sites only.** `api/utils/sites.js`: `knownSites()` (site profiles, the
  default site, and clean site codes on machines); Admin → Users only accepts those
  (`checkAllowedLocations`), and `GET /api/auth/sites` fills the pick-list. Lists
  saved as text (`'["workshop"]'`) are opened up wherever they are read.
- **One-time tidy-up (October 2026).** At start-up `repairSiteDataOnce()` copies the
  database (`ironlog_before_site_repair_*.db` in the copies folder), then moves
  records whose site was saved as a list to that site if it is real, else the
  default site; gives machines with no site the default site; and drops allowed
  locations that are not sites. Each change is in the audit log (module `sites`,
  user `system`). It runs once per database (`site_repairs` table).
- **Record guard.** `api/utils/siteGuard.js` runs before every route (server.js,
  preHandler). For a person limited to some sites it looks up any record the address
  names (work order, breakdown, inspection, purchase order, task, machine code…) and
  any machine, work order or breakdown named in the body, and answers 404 when that
  record is on another site. New record-number routes: add a line to `RECORD_RULES`.
- **Lists and totals** filter by the chosen site with `andSite(alias, req)` /
  `andAssetSite(column, req)` (the site code is checked before it reaches SQL), or
  `requireRequestSite(req, reply, site)` for one machine.
- **New records follow their machine.** Database triggers (`ensureSiteFollowsMachine`,
  run at start-up) give every new breakdown, work order, fuel/oil/hours log and
  inspection its machine's site.
- **Guard test.** `api/test/siteScope.test.js` fails when a route reads or changes site
  records without choosing a site. Handlers that did so when it was added are listed
  in `api/test/siteScopeBaseline.json`; that list was emptied in October 2026, so every
  route now chooses a site, has the record guard, or says why it needs neither. A handler that needs
  no filter says why in a `// site:` comment (a person's own records, or a record
  number the record guard checks). `api/test/siteIsolation.test.js` signs in a person
  limited to one of two sites and checks lists, record numbers and changes.

## Deliberately company-wide
- **Stock** (parts, stores, bins, stock levels): stores are not split by site yet.
- **Finance journals and integration exports**, **service templates**, and a
  person's **own** shifts, week and task inbox.
- **QR pre-start forms**: an operator scans one machine's code without signing in,
  so the machine code picks the machine on any site.
- **Uploads** only match machines on the chosen site; new machines from the machine
  upload are created on that site.

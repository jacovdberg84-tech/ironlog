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
- **Guard.** `api/test/siteScope.test.js` fails when a route reads or changes site
  records without choosing a site. Handlers that did so when it was added are listed
  in `api/test/siteScopeBaseline.json`; that list may only shrink.

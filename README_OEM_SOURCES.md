# Approved OEM link comparison

In Borris, expand **Check an approved OEM source**. Paste a public HTTPS page/PDF link and a question identifying the model and component. Ironlog retrieves that one reference and shows matching passages beside indexed internal manual excerpts. A loopback-configured LLM can draft a cited comparison; otherwise the evidence remains available with “comparison not established.” No external source is automatically trusted, saved as an approved library entry, or converted into a maintenance rule.

Approved exact hosts (verified from official sites):
- Bell: https://www.bellequipment.com/ and bare bellequipment.com
- Caterpillar: https://www.caterpillar.com/, https://www.cat.com/ and their bare hosts
- Atlas Copco: https://www.atlascopco.com/ and bare atlascopco.com
- Mercedes-Benz Trucks: https://www.mercedes-benz-trucks.com/ and bare mercedes-benz-trucks.com

Other subdomains/CDNs are not implicitly approved. Host additions require a reviewed code change. All redirects repeat allowlist and public-IP validation. DNS is resolved and the approved public address is pinned to the TLS request. URLs with credentials, queries, custom ports or non-HTTPS schemes are rejected. The fetcher sends no user cookies, authentication, internal records or question text. Only the user-supplied URL is fetched. Linked pages/assets/scripts are not executed or followed. This is not general internet search.

Limits: 15 seconds total network retrieval, at most four redirects, 25 MB response, PDF/HTML/plain text only. Poppler reads up to the first 100 physical PDF pages with a 20-second deadline. Online scans/diagrams need uploading and indexing through Workshop Library instead. One request runs at a time per API process. Local model comparison is limited to 15 seconds. References show retrieval time and source URL; revision is explicitly unverified. A public OEM page may still be marketing material, outdated or inapplicable to a particular serial number. Verify the original technical document before acting.

Deploy API and web changes together, install locked dependencies (`npm ci --omit=dev` in api), restart API, refresh browser. No new OS tools beyond existing Poppler. New direct packages: htmlparser2 10.0.0 and ipaddr.js 2.2.0. Authentication uses the existing Ironlog API hook; local authentication-disabled deployments retain their existing local access policy.

Validation: 37 automated tests passed, covering exact domains, private IP rejection, pinned DNS, redirect revalidation, text parsing, and existing workflows. A read-only live smoke test successfully retrieved Bell's official homepage. Production deployment and specific OEM bulletin/PDF checks have not yet been verified. Login-only portals, JavaScript-rendered pages, cross-domain downloads and signed links may require manual upload.

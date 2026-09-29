# Borris read-only stock chat

Ask Borris "Do we have hub seals in stock?", "Stock HS-12345", or "What is the unit price of HS-12345?". Explicit parts/manual requests append stock results to the cited manual answer.

Matches are labelled as exact stock codes supplied in the question, exact catalogue codes found in retrieved manual excerpts (fitment unverified), or description candidates. Manual code matching uses source excerpts rather than model-generated numbers. It does not maintain OEM cross-reference/alternative-part mappings, infer assembly fitment, or guarantee that a referenced part is the requested component. Multiple candidates require user selection of a stock code. Vague follow-ups such as "do we have it?" should repeat the component/code.

Balances sum stock_movements by part/location/bin, including negative entries and unspecified locations. Scope is all recorded stock locations, matching the existing stock depth data scope. Balances are on hand, not net availability after reservations. Prices come from parts.unit_cost; missing/zero values are shown as not recorded, and currency is not assumed. Catalogue cost is not a supplier quote. No stock adjustments, orders or maintenance writes occur.

Deploy the updated API and restart it. No schema migration, dependency install or document re-indexing is required. Tests cover exact and ambiguous matches, per-bin arithmetic, absent prices, citation retention and absence of database writes. Production stock balances still need verification after deployment.

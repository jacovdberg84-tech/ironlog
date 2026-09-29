# Stock categories

Every store item belongs to one category:

| Category | What goes in it |
|---|---|
| Oils & lubricants | Oils, greases and fluids (for example the Fuchs range) |
| G.E.T | Ground engaging tools: tips, adapters, retainers, cutting edges, grader blades |
| Parts | Workshop spares and service kits: seals, filters, bearings, hoses, bolts |
| Components | Major assemblies: engines, transmissions, pumps, cylinders and rams, hubs, radiators, turbos |
| Tyres | Tyres |

## How an item gets its category

- **Automatically.** When an item is created, IronLog suggests a category from its code and description. For example, codes starting with `MLF` (Fuchs) become Oils, "23.5R25" becomes Tyres, and "Tip Cat J300" becomes G.E.T. An item mentioning "oil" is not automatically an oil: an oil seal or oil filter stays in Parts.
- **By hand.** In Stock Control, each item card has a category picker. A category picked by hand is never changed automatically. To hand an item back to the automatic rules, pick its automatic category again, or call `POST /api/stock/part-category` with `category: "auto"`.
- **By CSV.** The parts CSV import accepts an optional `category` column (`oil`, `get`, `part`, `component`, `tyre`, or the names above).

## Where it is used

- The monthly and weekly stock report has one sheet per category and a "Stock by category" table on the summary.
- The Oils screen (formerly Lubricants), oil stock on hand, oil minimums and month-end oil stock only include Oils & lubricants items.
- Work order, maintenance and report cost splits between "oil" and "parts" use the same category.

The rules live in `api/utils/stockCategory.js`, and `api/test/stockCategory.test.js` checks them against real item names.

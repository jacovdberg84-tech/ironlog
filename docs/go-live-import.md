# Go-live import

Uploads → **Go-live import** loads a site's starting data from one Excel workbook (`api/utils/goLiveImport.js`, routes in `api/routes/upload/go-live.routes.js`).

- **Sheets:** Machines, Parts, Stock on hand, Suppliers, Components. The "Read me" sheet and the note on each heading explain the columns. Any sheet can be left empty.
- **Download** the workbook empty, or filled with the site's current data to correct and upload again.
- **Check** changes nothing. It lists every problem by sheet, row and column, warnings to look at, and how many records are new, changed or the same.
- **Import** runs the same check and applies the file in one transaction, only when there are no problems. Each import is kept under "Imports done" and in the audit log.
- **Keys:** machine code, part code, supplier code, and machine + component + position. Uploading the same file twice changes nothing.
- **Blank cells** clear an optional detail. Meter hours and stock are the exception: blank means leave as is.
- **Stock on hand** sets the balance. IronLog posts the difference as a `go_live_opening` movement and warns when stock has already moved.
- **Machines** go to the Site column's site, or the site chosen at the top, and only to sites the person may use. A machine already on another site is refused (move it on the Machines screen). Meter hours lower than the last daily reading are refused.
- **Who:** admin, supervisor and plant admin (workshop_admin).

The go-live checklist under the import shows, for the chosen site, what is in place: machines, make/model, meters, store items, opening stock, suppliers, components, people and daily hours.

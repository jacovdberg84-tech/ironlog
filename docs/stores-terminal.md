# Stores terminal (kiosk)

A touch screen at the stores counter, for example a Panasonic Toughbook, running one page:

`https://ironlog.ironlogafrica.com/web/stores-terminal.html`

Stock Control also links to it ("Stores terminal ↗" next to the tabs).

## Who uses it

| Person | Signs in with | Can do |
|---|---|---|
| Storeman (`storeman` / `stores`), plant admin (`workshop_admin`), admin, supervisor | Name tile + PIN | Issue parts, Receive delivery, Workshop requests, Find stock, Count stock |
| Technician (`artisan`) | Name tile + PIN | Collect parts for **their own** open work orders (lead or helper), Find stock |

- Set each person's 4–6 digit PIN in **User Admin**. Only people with a PIN appear on the terminal.
- Technicians are signed out after 90 seconds without a touch and straight after collecting. Storemen are signed out after 10 minutes. A yellow bar counts down the last 15 seconds.
- Every issue is written to the work order (or machine) under the signed-in person's name, the same as an issue from Stock Control. Issuing a workshop request marks that request received.
- Counts are blind: the terminal does not show the system quantity. A difference goes to a supervisor for approval, as in Counts & corrections.
- New parts, prices and bins are still set up in Stock Control on the office PC.
- EN/PT switch in the top bar.

## Scanning

The terminal reads whatever it is given, on any screen:

| Code | On the home screen | While issuing | While receiving / counting |
|---|---|---|---|
| Part label (code or the store QR) | Shows stock and bin | Adds the part | Adds / counts the part |
| Machine QR or code | Opens Issue with that machine's open job | Picks that machine's job | — |
| Work order number (`WO 123`) | Opens Issue for that job | Picks the job | — |

### Supplier barcodes on the boxes

Boxes arrive with the maker's barcode (EAN/UPC on filters, oils, bearings…). Link each one to its IronLog part once:

- **Scan it anywhere** on the terminal while a storeman is signed in. An unknown barcode opens *New barcode*: search the part, tap it. The scan then carries on (added to the issue or delivery, or stock shown).
- **Or from Find stock**: open the part, tap **＋ Link a box barcode**, scan the box.

After that, anyone scanning that box gets the part. A part can have several barcodes (different brands or pack sizes); each barcode belongs to one part, and moving it to another part asks first. Wrong links are removed with ✕ under *Box barcodes* on Find stock. Technicians who scan an unlinked barcode are told to ask the storeman. Linked barcodes also work in the Stock Control part search.

### IronLog's own labels (parts without a barcode)

For parts that come without a maker's barcode, print IronLog's label. It carries the part code as a **Code 128 barcode** (any scanner reads it) and a **QR** (phones), with the name and bin. No linking needed: the barcode *is* the part code.

- **Office PC:** Stock Control → **Setup** → **Part labels**. Add parts by search, or in one go: *All in stock without a box barcode* / *Received in the last 7 days*. Set the number of labels per part, choose the label size, **Print labels**. Clear the list once they printed correctly.
- **Terminal:** Find stock → part → **🖨 Label** → choose how many → *Add to the print list* (printed from the office PC), or *Print here* if a printer is connected to the terminal.
- **Sizes:** A4 sheets with 21 (63.5 × 38.1 mm, Avery L7160/J8160), 24 (70 × 37 mm) or 14 (99.1 × 38.1 mm, L7163) labels; label printers 62 × 29 mm (Brother DK-11209), 50 × 25 mm or 100 × 50 mm (Zebra). For a part-used A4 sheet set *Start at label no.*
- Print at **100 % / actual size** (turn off "fit to page"). For a label printer, choose its label size as the paper size in the printer settings. Use durable (polyester) labels in the workshop if possible.

**USB or Bluetooth barcode scanner:** plug it in, no setup. It types like a keyboard and ends with Enter; the terminal picks it up on any screen. Set the scanner to send Enter (CR) after each code (the default on most scanners).

**Phone as the scanner (until the scanner arrives):**

1. On the terminal tap **📱 Phone scanner**. A QR code appears.
2. Scan it with the phone's camera. The phone opens the scanner page.
3. Tap **Start camera** and allow the camera. Leave the page open.
4. Point the phone at labels. Each code shows on the terminal within a couple of seconds (the footer says "Phone scanner paired").

Android Chrome reads codes natively; iPhones use the bundled reader. If the camera will not start, type the code on the phone and tap Send. **New pairing** on the terminal disconnects earlier phones.

## Setting up the Toughbook (Windows 10/11)

### 1. A local kiosk account with Edge in kiosk mode (recommended)

1. **Settings → Accounts → Other users → Set up a kiosk** (Windows 11: *Accounts → Other users → Kiosk*).
2. Name it `Stores`, choose **Microsoft Edge**, then **As a digital sign or interactive display** (full screen, one site, no address bar). Do not pick "As a public browser".
3. URL: `https://ironlog.ironlogafrica.com/web/stores-terminal.html`
4. Idle restart: **Never** (the terminal signs people out by itself).
5. Restart. Windows signs in to the kiosk account and opens the terminal full screen. Ctrl+Alt+Del gets you back to an admin account.

### 2. Or: an Edge shortcut that starts with Windows

For a quicker setup on a normal account, create a shortcut with this target and put it in `shell:startup`:

```
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --kiosk https://ironlog.ironlogafrica.com/web/stores-terminal.html --edge-kiosk-type=fullscreen --no-first-run
```

Alt+F4 leaves the kiosk.

### 3. Touch and power settings

- **Tablet mode / touch keyboard:** Settings → Time & language → Typing → Touch keyboard → *Show the touch keyboard when there's no keyboard attached* = On. The keyboard appears for the invoice number and search boxes.
- **Power:** Settings → System → Power → Screen and sleep → *Never* when plugged in.
- **Updates:** set active hours to cover the workshop shift, so Windows does not restart mid-shift.
- **Display:** 125–150 % scaling suits the Toughbook screen; the terminal adjusts to any size.
- **Edge:** allow sound for the site (scan beeps). Turn off "Offer to save passwords".

### 4. Scanner checks

Open Notepad, scan a part label: the code should appear followed by a new line. If not, set the scanner to "USB HID keyboard" mode with a CR suffix (see its quick-start sheet).

## For IT

- Pages: `web/stores-terminal.html` / `.js` / `.css`, phone page `web/store-scanner.html`.
- API: `api/routes/stock/terminal.routes.js` under `/api/stock/terminal/*` (home, work-orders, requests, lookup, issue, barcodes). Box barcodes live in `part_barcodes` (`api/utils/partBarcodes.js`). Receive and count reuse `/api/stock/deliveries` and `/api/stock/cycle-count`.
- PIN sign-in: `GET /api/auth/pin/roster?terminal=stores` also lists stores staff; `POST /api/auth/pin/login` accepts `storeman`, `stores` and `workshop_admin` besides technicians.
- Phone relay: `POST /api/stock/terminal/scan` and `GET /api/stock/terminal/scans` are public. The terminal makes a random 128-bit key and shows it only in the pairing QR; scans are held in memory for two minutes and only codes pass through. A server restart drops waiting scans, nothing else.
- Labels: `web/part-labels.js` (shared printer), `web/app/stock-labels.js` (Stock Control card), `api/routes/stock/labels.routes.js` (`/api/stock/labels/queue`, `/api/stock/labels/parts`), table `part_label_queue`.
- Bundled libraries (MIT): `web/vendor/qrcode-generator-1.4.4.js` (pairing and label QR), `web/vendor/jsbarcode-code128-3.11.6.min.js` (label barcodes) and `web/vendor/zxing-browser-0.1.5.min.js` (camera reading where the browser has no BarcodeDetector).

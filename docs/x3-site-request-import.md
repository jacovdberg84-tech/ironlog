# Sage X3 Site Requests from IronLog

IronLog writes requisitions as an import file for AML's custom Site Request screen (window OZS1). The field codes come from an X3 export of Site Request VS120-00353.

## For the X3 administrator: the import template
Create an import template (Usage → Imports/exports → Import/export templates) on the Site Request object:

- **File:** delimited text, field separator `;`, field delimiter `"`, dates `DD/MM/YYYY` (all four can be changed in IronLog to match the template).
- **Line types:** `E` = Site Request header (one per requisition), `L` = Site Request Detail line (one per part), each `L` belongs to the `E` above it.
- **Header (E) fields, in order:** ZSOURCESA, ZDELFCY, ZFCY, ZREF, ZBILLCUR, ZCONTRACT, ZDLVPIO, ZREQTYP, ZSIGNCHNL, ZREQUESTBY, ZBUYER, ZDATEREQ, ZHBPSNUM, ZSUPPQUOTE, ZPREFSUPP, ZBPCORD, ZBPCINV, ZBPAADD, ZCUR, ZAX1, ZVIN, ZODOUNIT, ZODOREADING, ZAX2, ZAX3, ZSITEJOB
- **Line (L) fields, in order:** ZITMREF, ZOEM, ZPRINTDESC, ZUOM, ZQTY, ZEQUIPMENT, ZODOUNIT, ZODOREADING, ZJOBCARD, ZPROJECT, ZBPSNUM, ZGROPRI, ZLCUR, ZDISTYP, ZDISCRGVAL, ZNETPRI, ZDATEREQ

X3 assigns the Site Request number (ZSREQNO); IronLog does not send it. ZREF (Manual Req no.) carries IronLog's requisition number, e.g. `IL000123`, so the two can be matched. ZSOURCESA is sent as the X3 menu number (2 = Yes). Display-only fields (ZZ…), status, posted and line totals are not sent; X3 works them out.

The full table, with where IronLog takes each value from, is in IronLog under Procurement → Sage X3 Site Requests → File layout.

## For buyers and stores
1. Procurement → **Sage X3 Site Requests**: fill in the X3 codes once (sites, request type, signature channel, buyer, customer, project, the X3 product used for parts) and the people and supplier code lists. Suppliers in IronLog's supplier list use their IronLog supplier code.
2. On a finalised requisition press **Download for X3**, or **Download all waiting for X3**. IronLog lists anything missing first (codes, drafts, quantities) and warns about unknown suppliers, missing machines and requisitions already sent.
3. Import the file in X3.
4. Press **X3 number** on the requisition and type the number X3 gave it (e.g. VS120-00353).

Settings: admin, supervisor, procurement. Download and X3 number: admin, supervisor, procurement, stores. All are audited.

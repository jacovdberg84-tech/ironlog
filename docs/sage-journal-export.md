# Sage journal batches

IronLog's journal runs (Finance → Summarized Journal Runs) can be downloaded as a batch for Sage.

## Set up once (Finance → Sage journal batches)
- **Sage account** for each IronLog account. Only the accounts a run actually uses must be filled in; the download says which are missing.
- **Send to Sage**: which costs go. Downtime is a notional cost and is off by default.
- **Cost centre**: part of the account (`5200` + `-` + `PIT` = `5200-PIT`, the usual Sage 300 set-up), its own column (CSV only), or left out. Lines without a cost centre use "Cost centre when none is set".
- **Layout**:
  - *Sage 300 import sheet*: Excel with `Journal_Headers` and `Journal_Details` sheets for GL → Import Journal Entries. Amounts are positive for debits and negative for credits. Sage gives the batch its own number on import.
  - *Journal lines (CSV)*: `Date, Reference, Account, [Cost centre], Description, Debit, Credit`, for Sage Evolution's journal batch import or another system.
- **Lines**: one per account (and cost centre), or one per account and machine.

## Each month
1. Build the run, pick it, press **Check for Sage** to see the batch.
2. **Download for Sage**. The batch is dated at the run's end date, its reference is the prefix plus the run number, and it only downloads when debits equal credits (rounding of a few cents goes on the largest line).
3. Import in Sage, then **Mark Exported** / **Mark Posted** with Sage's batch number.

Before the first live month, import one batch into a test company and compare it with your Sage template; column names can be matched in the import screen.

Settings: admin, supervisor, finance. Download: admin, supervisor, procurement, finance, plant manager. Both are audited.

# Internal Workshop Library

Open Workshop Library in Ironlog. Upload PDF manuals, bulletins, technical documents or workshop fixes, with a title, document type, manufacturer, model/engine, revision and serial applicability. Search these fields and download the original PDF. Files are limited to 100 MB each; each upload remains a separate record so earlier editions are retained.

Uploads require admin, supervisor, workshop_admin or plant_manager roles. Reads and downloads use the existing Ironlog API authentication policy. In deployments where authentication is disabled, that existing local-access policy still applies. Uploaded material is a reference, not approved engineering guidance.

The files live in `workshop-files` beneath `IRONLOG_DATA_DIR` (or the API working directory). This is outside the public uploads path. Back up this directory alongside the SQLite database; existing database-only backups do not contain the PDFs. Keep the directory persistent across deployments and ensure the API user can write it. The document table is created on API startup. No new package dependencies are required. Configure any reverse proxy to accept 100 MB uploads if that size is needed.

Deploy API and web changes together, restart the API, and refresh the page. The previous external library remains accessible through an explicit button; selecting Workshop Library no longer opens it automatically.

This release stores original PDFs and metadata and supports explicit local text/OCR indexing and cited library answers. Repair-outcome learning is not yet enabled.

## Borris indexing and cited answers

After deployment, select **Index for Borris** on each existing document. Use **Search / refresh** to check progress. New uploads also need this explicit indexing step. Jobs run serially; interrupted jobs can be retried after a server restart. Originals are preserved. Indexes and job state live in SQLite and contain document text, so protect database backups accordingly.

The API host must provide Poppler (`pdftotext`, `pdftoppm`) and Tesseract with English language data. Install these through the server operating system's package manager. On Debian/Ubuntu the packages are `poppler-utils`, `tesseract-ocr`, and `tesseract-ocr-eng`. Windows installations can supply absolute executable paths through `PDFTOTEXT_BIN`, `PDFTOPPM_BIN`, and `TESSERACT_BIN`. Optional `WORKSHOP_OCR_LANG` defaults to `eng`. Restart the API after changing its environment. Missing Poppler causes a failed index with an actionable message; missing OCR leaves scanned pages flagged as partial.

Ask questions in **Ask the workshop manuals**. The ordinary Borris chat also routes questions containing manual, bulletin, fault code, torque or specification to the indexed library. Citations use physical PDF page numbers, which may differ from the printed page labels. Search currently uses lexical full-text matching, not semantic embeddings. Include the machine and specific component/fault code to improve results. Document-specific API queries accept `document_id` on `/api/workshop/ask`.

Only a loopback-configured model endpoint is used for manual answer synthesis. If no local model is configured or it fails, the response shows retrieved excerpts and source pages. Remote model endpoints do not receive manual excerpts through this feature. Page references identify retrieval evidence; generated wording is not independently verified. Always check the original diagram/procedure, edition, serial applicability and OCR values before acting. No operational records or approved engineering rules are changed.

Pages with fewer than 40 non-whitespace characters receive OCR. Mixed text-and-image pages may still contain unindexed labels in diagrams. Sparse pages, extraction limits and OCR failures are flagged. Documents over 2000 pages must be split. Indexing does not yet interpret diagrams, approve procedures or learn from repair outcomes.

Validation: 33 automated tests pass, including extraction/OCR command handling with simulated tool output, physical-page preservation, SQLite full-text retrieval and citations. Real extraction/OCR on the production manuals and production executable installation remain deployment verification steps.


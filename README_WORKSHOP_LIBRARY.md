# Internal Workshop Library

Open Workshop Library in Ironlog. Upload PDF manuals, bulletins, technical documents or workshop fixes, with a title, document type, manufacturer, model/engine, revision and serial applicability. Search these fields and download the original PDF. Files are limited to 100 MB each; each upload remains a separate record so earlier editions are retained.

Uploads require admin, supervisor, workshop_admin or plant_manager roles. Reads and downloads use the existing Ironlog API authentication policy. In deployments where authentication is disabled, that existing local-access policy still applies. Uploaded material is a reference, not approved engineering guidance.

The files live in `workshop-files` beneath `IRONLOG_DATA_DIR` (or the API working directory). This is outside the public uploads path. Back up this directory alongside the SQLite database; existing database-only backups do not contain the PDFs. Keep the directory persistent across deployments and ensure the API user can write it. The document table is created on API startup. No new package dependencies are required. Configure any reverse proxy to accept 100 MB uploads if that size is needed.

Deploy API and web changes together, restart the API, and refresh the page. The previous external library remains accessible through an explicit button; selecting Workshop Library no longer opens it automatically.

This release stores and retrieves original PDFs and searchable metadata. PDF text extraction, OCR, Borris citations and repair-outcome learning are not enabled yet. No document contents are sent to an external model by these upload routes.

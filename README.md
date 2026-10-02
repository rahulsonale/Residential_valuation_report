# Residential Valuation Report Generator

## Run the Express backend

From this project folder, install dependencies once with `npm install`, then start the backend with:

```powershell
npm run backend
```

By default it listens at `http://127.0.0.1:3000`. Set `BACKEND_HOST` or `BACKEND_PORT` to override those defaults.

## Draft upload flow

The backend accepts one `.docx` file up to 10 MB. The single-call endpoint is:

```text
POST /api/document/process
multipart/form-data field: file
```

For each upload it extracts a structure index, saves an internal JSON snapshot under `output/layouts/<layoutId>.json`, generates a DOCX, and returns that DOCX as the response. The response includes `X-Layout-Id` and `X-Layout-Json` headers so the caller can retrieve the saved JSON with `GET /api/layout/<layoutId>`.

The steps can also be called separately:

- `POST /api/layout/extract` with multipart field `file` saves and returns the JSON snapshot.
- `GET /api/layout/<layoutId>` reads a saved snapshot.
- `POST /api/layout/generate` with the snapshot JSON in the request body returns a DOCX.
- `GET /api/health` checks whether the backend is running.

Each upload gets a unique layout ID, so concurrent uploads do not overwrite each other's JSON. Snapshots remain in `output/layouts` for subsequent handling; generated DOCX response files are removed after download.

## Current automation scope

The snapshot contains a readable index of sections, page sizes/orientations, tables, cells, spans, border declarations, and image counts, plus the source DOCX package encoded as base64 with a SHA-256 integrity check. Generation restores that package exactly, which preserves the whole report structure and current text/values, including content that the structure index does not interpret.

This gives a complete upload → internal JSON → generated DOCX round-trip for the supplied draft. It does not yet infer or change report values, fill a blank template from another source, or automatically correct the header. Those transformations need defined input fields and mapping rules; header correction can remain manual as agreed.

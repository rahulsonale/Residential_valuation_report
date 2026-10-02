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


# Angular Frontend: DOCX to JSON to DOCX

## Purpose

The Angular frontend provides a browser interface for the existing Express backend. It lets a user:

1. Choose a reference `.docx` file.
2. Send it to Express to extract its layout as JSON.
3. View the JSON returned by Express in the page.
4. Send that JSON to Express to generate and download a new `.docx` file.

The backend remains responsible for reading and writing Word documents. Angular provides the user interface and sends the API requests.

## Project layout

The Angular application is in:

```text
D:\Residential-Valuation-Report\frontend
```

Important frontend files:

| File | Purpose |
| --- | --- |
| `src/app/app.html` | Page structure, file picker, buttons, status messages, and JSON preview. |
| `src/app/app.ts` | Frontend state and API request logic. |
| `src/app/app.css` | Styles for the page and JSON preview. |
| `src/app/app.config.ts` | Registers Angular services, including `HttpClient`. |
| `src/proxy.conf.json` | Forwards frontend `/api/...` requests to Express during development. |
| `angular.json` | Angular CLI configuration; points the development server to the proxy file. |

The Express backend files remain in the project root's `backend` and `generators` folders.

## 1. Create and install the Angular application

From PowerShell:

```powershell
cd "D:\Residential-Valuation-Report"
npx @angular/cli@21 new frontend --routing=false --style=css
```

Choose **No** for analytics if you do not want to share anonymous usage information. Choose **No** for SSR/SSG; this first version is a browser-based interface. Leave optional AI instruction-file choices unselected.

If Angular creates the frontend files but dependencies are not installed, enter the frontend folder and install them:

```powershell
cd "D:\Residential-Valuation-Report\frontend"
npm install --legacy-peer-deps
```

The `--legacy-peer-deps` option was used after npm 10.9.3 returned an `edgesOut` dependency-resolution error during a normal install.

## 2. Configure the development proxy

Angular runs on port `4200`; Express runs on port `3000`. The proxy forwards `/api/...` requests from Angular to Express so the browser can use relative API paths.

Create `frontend/src/proxy.conf.json`:

```json
{
  "/api/**": {
    "target": "http://127.0.0.1:3000",
    "secure": false
  }
}
```

In `frontend/angular.json`, add the `options` block to the `serve` configuration:

```json
"serve": {
  "builder": "@angular/build:dev-server",
  "options": {
    "proxyConfig": "src/proxy.conf.json"
  },
  "configurations": {
```

Keep the existing configuration entries after this block. Restart `ng serve` after changing the proxy file or Angular configuration.

## 3. Enable Angular HTTP requests

In `frontend/src/app/app.config.ts`, register `provideHttpClient()`:

```ts
import {
  ApplicationConfig,
  provideBrowserGlobalErrorListeners,
} from '@angular/core';
import { provideHttpClient } from '@angular/common/http';

export const appConfig: ApplicationConfig = {
  providers: [
    provideBrowserGlobalErrorListeners(),
    provideHttpClient(),
  ],
};
```

`HttpClient` is the Angular service used to send requests to the Express API.

## 4. What `app.ts` does

`frontend/src/app/app.ts` contains the frontend behavior:

- Keeps the selected file, extracted layout, loading state, and error text.
- Sends the selected file as `multipart/form-data` to `POST /api/layout/extract`.
- Stores the returned layout in the UI state.
- Sends the layout JSON to `POST /api/layout/generate`.
- Receives the generated DOCX as a `Blob` and starts a browser download.

The component uses Angular signals for UI state. Angular 21 enables zoneless change detection by default, and updating template-read signals notifies Angular to refresh the page after an asynchronous API response.

The imports for the component should be separated correctly:

```ts
import { Component, inject, signal } from '@angular/core';
import { JsonPipe } from '@angular/common';
import { HttpClient } from '@angular/common/http';
```

`JsonPipe` comes from `@angular/common`; `HttpClient` comes from `@angular/common/http`.

## 5. What `app.html` does

`frontend/src/app/app.html` displays the workflow:

- File picker accepts `.docx` documents.
- **Extract layout JSON** starts extraction and shows the loading state.
- On success, the page displays the layout ID, section/table counts, and a collapsible JSON preview.
- **Generate DOCX** becomes available after a layout has been extracted.
- Generation status or an error message appears below the button.

The JSON preview uses Angular's `JsonPipe`. A complete extracted layout can be large because the backend JSON includes Word package data; expand the preview only when needed.

## 6. Run the application

Use two terminal windows in VS Code.

**Terminal 1 — Express backend**, from the project root:

```powershell
cd "D:\Residential-Valuation-Report"
npm run backend
```

The backend should report that it is listening at `http://127.0.0.1:3000`.

**Terminal 2 — Angular frontend**, from the frontend folder:

```powershell
cd "D:\Residential-Valuation-Report\frontend"
npm start
```

Open the Angular page at:

```text
http://localhost:4200
```

To check the proxy and backend health, visit:

```text
http://localhost:4200/api/health
```

A healthy backend returns JSON with `ok: true` and `status: "ready"`.

## 7. Extract JSON through the interface

1. Open `http://localhost:4200`.
2. Choose a reference `.docx` file.
3. Click **Extract layout JSON**.
4. Wait for the loading label to finish.
5. Review the layout ID and section/table counts.
6. Expand **View extracted JSON** to inspect the response.

The Angular request is sent to `/api/layout/extract`. Express extracts the layout and saves a copy under:

```text
D:\Residential-Valuation-Report\output\layouts
```

The exact filename uses the generated layout ID. The API also returns the JSON to Angular for display.

## 8. Generate a DOCX through the interface

1. After extraction succeeds, click **Generate DOCX**.
2. Angular posts the extracted JSON to `/api/layout/generate`.
3. Express passes the layout to the Word generator and returns the generated document.
4. Angular handles the response as binary data (`Blob`) and starts downloading `generated-layout.docx`.
5. Open the downloaded document in Word and compare it with the reference.

## Current progress

- Angular frontend scaffold created.
- Frontend dependencies installed using `npm install --legacy-peer-deps` after the npm resolver error.
- Angular dev server running.
- Proxy reaches the Express health endpoint.
- DOCX extraction from the Angular interface is working.
- Extracted JSON is visible in the UI.
- DOCX generation/download code has been added as the next frontend stage; successful generation from the UI still needs to be confirmed.

## Troubleshooting notes

### `npx ng serve` cannot find the executable

Install the Angular project dependencies from the frontend folder, then start the app:

```powershell
npm install --legacy-peer-deps
npm start
```

### npm reports `Cannot read properties of null (reading 'edgesOut')`

Retry installation from the frontend folder with:

```powershell
npm install --legacy-peer-deps
```

### Angular reports that `JsonPipe` is not exported by `@angular/common/http`

Use this import in `app.ts`:

```ts
import { JsonPipe } from '@angular/common';
```

Do not import `JsonPipe` from `@angular/common/http`.

### Browser shows `Route not found` for `/api/layout/extract`

The extraction endpoint is POST-only and requires a file upload. Use the Angular form or Postman with a multipart form-data field named `file`; opening the endpoint directly in a browser sends a GET request instead.

### API calls fail through the proxy

Confirm that Express is running in the other terminal, the proxy file is `frontend/src/proxy.conf.json`, and `angular.json` contains the `proxyConfig` option. Restart Angular after proxy configuration changes.

# UBI Umbare DOCX to JSON to DOCX Workflow

## Goal

Use the new reference document, `UBI - Umbare (3).docx`, to produce a JSON representation of its document structure, then use that JSON to generate a DOCX and compare the result with the reference. Image matching in tables is lower priority for this pass.

## Reference review

The reference was reviewed structurally before processing. It is a Union Bank of India valuation report with multiple sections and tables. The initial inspection found 3 portrait A4 sections, 12 main tables, and 268 body paragraphs. These are structural inspection results; a rendered page-by-page visual review was not completed at that point.

## Work completed so far

1. Chose the new reference document: `UBI - Umbare (3).docx`.
2. Confirmed that the existing Express backend is intended to support the same extract-then-generate workflow used for the previous reference. No code changes were planned for this first attempt, provided both endpoints work for this document.
3. In Postman, sent the reference DOCX to the extraction endpoint using `multipart/form-data`, with a file field named `file`.
4. Received a JSON response and saved it in `D:\Residential-Valuation-Report\output\layouts` as `ubi-umbare-layout.json` (or the chosen filename if renamed).
5. Prepared the next step: submit that saved JSON to the generation endpoint. The user has not yet confirmed that this request has been sent or that the generated DOCX has been saved.

## Step 1: Extract the reference to JSON

In Postman, create a request with:

- Method: `POST`
- URL: `http://127.0.0.1:3000/api/layout/extract`
- Body type: `form-data`
- Key: `file`
- Change the field type from **Text** to **File**, then select `UBI - Umbare (3).docx`.

Click **Send**. The successful response should be JSON describing the extracted document layout. Save the response as a new JSON file in:

```text
D:\Residential-Valuation-Report\output\layouts
```

Use a distinct filename, for example `ubi-umbare-layout.json`, so the previous reference's JSON is not overwritten.

## Step 2: Generate a DOCX from the JSON

In Postman, create another request:

- Method: `POST`
- URL: `http://127.0.0.1:3000/api/layout/generate`
- Body type: **raw**
- Raw format: **JSON**
- Body: the complete contents of the newly saved layout JSON file.

Click **Send**. A successful request should return status `200 OK` and a DOCX file as binary response data. DOCX is a ZIP-based format, so the response may look like unreadable characters if Postman displays it as text; that is expected for binary data.

Save the response as a `.docx` file, for example `ubi-umbare-generated.docx`. Ensure the saved filename ends in `.docx`, not `.json` or `.txt`.

## Step 3: Compare the generated DOCX with the reference

Open both documents in Word and review:

- Page count and page flow
- Section/page orientation and margins
- Table count, placement, row and column structure
- Text and cell content
- Borders, shading, fonts, and alignment
- Images inside tables (deferred/lower priority for this pass)

Record any differences. If extraction or generation fails, or the comparison shows missing structure or formatting, inspect the relevant endpoint/code and make a targeted change before repeating the workflow.

## Current status

- JSON extraction: **completed**; JSON saved in the layouts folder.
- DOCX generation from that JSON: **next step; completion not yet confirmed**.
- Reference comparison: **pending until the generated DOCX is saved**.

## Backend notes

The backend should be running locally on port `3000` for these Postman requests. The endpoints documented here are the two-stage workflow:

- `POST /api/layout/extract` — DOCX upload to JSON response.
- `POST /api/layout/generate` — JSON request to generated DOCX response.

Opening `/api/layout/extract` directly in a browser is not the same as sending a file upload. Also, a browser `GET` to a `POST`-only endpoint can return `Route not found`; use Postman with the method and body described above.

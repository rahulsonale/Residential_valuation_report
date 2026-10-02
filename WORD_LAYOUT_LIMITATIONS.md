# Word Layout Limitations and Design Rules

## Purpose

This file records how Microsoft Word may handle the report layout. The JSON should describe the structure while allowing Word to position content safely on the page.

## Page size

- The generated report must use A4 portrait paper.
- A4 is 210 mm × 297 mm, or approximately 595.28 × 841.89 points.
- The usable area is smaller than the paper because of the margins and the printer's printable area.
- The final margins still need to be agreed and recorded in the JSON.

## Margins, header, and footer

- Keep report content inside the page margins.
- Leave enough room for the header and footer so they do not overlap the main content.
- Printers may have different non-printable edge areas, so content placed close to a page edge may be clipped.
- Do not use exact X/Y positions for normal text and tables unless a specific element must be anchored.

## Text and page flow

- Word wraps text when it reaches the edge of its cell or paragraph area.
- Longer text makes rows and sections taller.
- When content no longer fits on a page, Word continues it on a later page.
- The number of pages may change when text, fonts, margins, or printer settings change.

## Tables

- Tables can continue onto later pages.
- A table's header row can be set to repeat on later pages.
- Word may split a row across pages unless the row is configured to stay together.
- If a row is kept together but cannot fit in the remaining space, Word moves it to the next page.
- Merged cells and nested tables can make table sizing and page breaks harder to control.
- Avoid fixed row heights when cells contain text that may wrap.

## Blank cover table and border overrides

- Build the blank cover as one 18-row, two-column table. Keep its fixed 9330-twip width, 553-twip indent, 4861/4469-twip column split, row minimum heights, and beige cell fills in `config/report-layout-v2.json`.
- Apply the black single-line outer and inside borders at table level, then retain each cell's `borders` overrides. A cell border set to `nil` suppresses only that edge; omitting an override leaves the table border visible. Do not replace these mixed per-cell settings with a uniform border preset.
- Keep the first cover row merged across both columns. The tall logo band below it hides the center divider, while later rows restore it and selectively suppress horizontal edges, matching the supplied Word template.
- Use separate Word sections for the cover and report details so their page margins can match the reference and the tall blank cover table stays on one page.

## Images and placeholders

- The QR section should be represented by an empty bordered box only.
- Do not include a QR image, QR data, or QR instruction text in the blank layout.
- The company logo and stamp need their image assets and placement rules before they can be reproduced accurately.
- Images may move when nearby text or tables grow, depending on how they are anchored.

## Layout choices for this project

- Use automatic text flow for regular paragraphs and tables.
- Use explicit page breaks only where the report clearly starts a new section.
- Use fixed positioning only for elements that must visually overlap or stay in a specific place.
- Check the generated Word document on A4 pages because JSON alone cannot guarantee the final printed appearance.

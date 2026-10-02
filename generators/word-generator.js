const fs = require("node:fs/promises");
const crypto = require("node:crypto");
const path = require("node:path");
const JSZip = require("jszip");
const {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  HeightRule,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} = require("docx");

const TABLE_BORDER_SIDES = {
  top: "top",
  bottom: "bottom",
  left: "left",
  right: "right",
  insideHorizontal: "insideHorizontal",
  insideVertical: "insideVertical",
};

function makeBorder(border) {
  if (
    border === "nil" ||
    border === "none" ||
    border === false ||
    (border &&
      typeof border === "object" &&
      ["nil", "none"].includes(border.style))
  ) {
    return { style: BorderStyle.NIL, size: 0, color: "FFFFFF" };
  }
  const style =
    typeof border === "string" ? border : (border.style ?? "single");
  return {
    style: BorderStyle[style.toUpperCase()] ?? BorderStyle.SINGLE,
    size: typeof border === "object" ? (border.size ?? 8) : 8,
    color: typeof border === "object" ? (border.color ?? "000000") : "000000",
  };
}

function makeBorders(borders, sideMap = TABLE_BORDER_SIDES) {
  if (!borders) return undefined;
  return Object.fromEntries(
    Object.entries(borders)
      .filter(([side]) => sideMap[side])
      .map(([side, border]) => [sideMap[side], makeBorder(border)]),
  );
}

function shadingColor(shading) {
  const color = typeof shading === "string" ? shading : shading.fill;
  return color.toLowerCase() === "beige" ? "C4BC96" : color;
}

function paragraph(text = "", options = {}) {
  return new Paragraph({
    children: [new TextRun(String(text ?? ""))],
    ...options,
  });
}

function makeCell(cell, widthDxa) {
  const isObject = cell !== null && typeof cell === "object";
  const children =
    isObject && cell.type === "table"
      ? [makeTable(cell, widthDxa)]
      : [
          paragraph(
            isObject ? (cell.text ?? cell.value ?? cell.role ?? "") : cell,
            {
              bold: Boolean(cell?.bold),
              alignment: {
                left: AlignmentType.LEFT,
                center: AlignmentType.CENTER,
                right: AlignmentType.RIGHT,
              }[cell?.alignment],
            },
          ),
        ];

  const cellOptions = { children };

  if (Number.isInteger(cell?.columnSpan)) {
    cellOptions.columnSpan = cell.columnSpan;
  }
  if (Number.isInteger(cell?.rowSpan)) {
    cellOptions.rowSpan = cell.rowSpan;
  }
  if (Number.isFinite(widthDxa)) {
    cellOptions.width = { size: widthDxa, type: WidthType.DXA };
  }
  if (isObject && cell.shading) {
    cellOptions.shading = {
      fill: shadingColor(cell.shading),
    };
  }
  if (isObject && cell.borders) {
    cellOptions.borders = makeBorders(cell.borders, {
      top: "top",
      bottom: "bottom",
      left: "left",
      right: "right",
      start: "start",
      end: "end",
    });
  }
  if (isObject && cell.verticalAlignment) {
    cellOptions.verticalAlign = cell.verticalAlignment;
  }
  if (isObject && cell.margins) {
    cellOptions.margins = cell.margins;
  }

  return new TableCell(cellOptions);
}

function makeTable(tableOrRows, totalWidthDxa = 9360) {
  if (Array.isArray(tableOrRows)) {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: tableOrRows.map(
        (cells) =>
          new TableRow({
            children: cells.map((cell) => makeCell(cell)),
          }),
      ),
    });
  }

  totalWidthDxa = tableOrRows.widthDxa ?? totalWidthDxa;
  const rows = tableOrRows.rows;
  const columnCount = Math.max(
    1,
    ...rows.map((row) =>
      row.cells.reduce((count, cell) => count + (cell.columnSpan ?? 1), 0),
    ),
  );

  const percentages =
    tableOrRows.columnWidthsPercent ??
    Array(columnCount).fill(100 / columnCount);

  const columnWidths = percentages.map((percent, index) =>
    index === percentages.length - 1
      ? totalWidthDxa -
        percentages
          .slice(0, -1)
          .reduce(
            (sum, earlier) => sum + Math.round((totalWidthDxa * earlier) / 100),
            0,
          )
      : Math.round((totalWidthDxa * percent) / 100),
  );

  return new Table({
    width: { size: totalWidthDxa, type: WidthType.DXA },
    columnWidths,
    ...(tableOrRows.indentTwips
      ? { indent: { size: tableOrRows.indentTwips, type: WidthType.DXA } }
      : {}),
    ...(tableOrRows.borders
      ? { borders: makeBorders(tableOrRows.borders) }
      : {}),
    ...(tableOrRows.margins ? { margins: tableOrRows.margins } : {}),
    rows: rows.map((row) => {
      let columnIndex = 0;

      return new TableRow({
        ...(row.heightTwips
          ? {
              height: {
                value: row.heightTwips,
                rule: HeightRule.ATLEAST,
              },
            }
          : {}),
        ...(row.cantSplit ? { cantSplit: true } : {}),
        children: row.cells.map((cell) => {
          const span = cell.columnSpan ?? 1;
          const cellWidth = columnWidths
            .slice(columnIndex, columnIndex + span)
            .reduce((sum, width) => sum + width, 0);

          columnIndex += span;
          return makeCell(cell, cellWidth);
        }),
      });
    }),
  });
}

function rolePlaceholder(role = "") {
  return String(role)
    .replace(/[-_]/g, " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function renderBlock(block) {
  const alignment = {
    left: AlignmentType.LEFT,
    center: AlignmentType.CENTER,
    right: AlignmentType.RIGHT,
  }[block.alignment];

  if (block.type === "table") {
    return makeTable(block);
  }

  if (block.type === "image") {
    return paragraph(`[${rolePlaceholder(block.role)}]`, { alignment });
  }

  if (block.type === "text") {
    const isHeading = block.role?.includes("heading");
    return paragraph(block.value ?? rolePlaceholder(block.role), {
      alignment,
      heading: isHeading ? HeadingLevel.HEADING_2 : undefined,
      pageBreakBefore: block.pageBreakBefore,
    });
  }

  return paragraph(`[${rolePlaceholder(block.type)}]`);
}

async function generateWordDocument({ reportData, reportLayout, outputPath }) {
  console.log(`[Generator] Creating DOCX: ${outputPath}`);
  if (
    reportLayout.packagePayload?.encoding === "docx-parts-json" &&
    Array.isArray(reportLayout.packagePayload.parts)
  ) {
    const zip = new JSZip();
    const requiredParts = new Set(["[Content_Types].xml", "word/document.xml"]);

    for (const part of reportLayout.packagePayload.parts) {
      if (
        !part ||
        typeof part.path !== "string" ||
        !part.path ||
        part.path.startsWith("/") ||
        part.path.split("/").includes("..") ||
        typeof part.content !== "string"
      ) {
        throw new Error("The JSON contains an invalid Word package part.");
      }

      const partBuffer =
        part.encoding === "utf8"
          ? Buffer.from(part.content, "utf8")
          : part.encoding === "base64"
            ? Buffer.from(part.content, "base64")
            : null;
      if (!partBuffer) {
        throw new Error(
          `Unsupported encoding for Word package part: ${part.path}`,
        );
      }
      if (part.sha256) {
        const partHash = crypto
          .createHash("sha256")
          .update(partBuffer)
          .digest("hex");
        if (partHash !== part.sha256) {
          throw new Error(
            `Word package part failed its integrity check: ${part.path}`,
          );
        }
      }

      requiredParts.delete(part.path);
      zip.file(part.path, partBuffer, {
        createFolders: false,
        ...(part.date ? { date: new Date(part.date) } : {}),
      });
    }

    if (requiredParts.size) {
      throw new Error(
        `The JSON is missing required Word parts: ${[...requiredParts].join(", ")}`,
      );
    }

    const packageBuffer = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    });
    await fs.mkdir(path.dirname(outputPath), { recursive: true });
    await fs.writeFile(outputPath, packageBuffer);
    return;
  }

  if (
    reportLayout.packagePayload?.encoding === "base64" &&
    typeof reportLayout.packagePayload.content === "string"
  ) {
    const packageBuffer = Buffer.from(
      reportLayout.packagePayload.content,
      "base64",
    );
    const expectedHash = reportLayout.packagePayload.sha256;
    if (expectedHash) {
      const actualHash = crypto
        .createHash("sha256")
        .update(packageBuffer)
        .digest("hex");
      if (actualHash !== expectedHash) {
        throw new Error(
          "The embedded Word package failed its integrity check.",
        );
      }
    }
    await fs.mkdir(path.dirname(outputPath), { recursive: true });
    await fs.writeFile(outputPath, packageBuffer);
    return;
  }

  const page = reportLayout.document.page;
  const defaultMargins = reportLayout.document.marginsPt ?? {};
  const sections = reportLayout.sections.map((section) => {
    const margins = { ...defaultMargins, ...(section.marginsPt ?? {}) };
    const children = section.blocks.flatMap((block) => {
      const rendered = renderBlock(block);
      return Array.isArray(rendered) ? rendered : [rendered];
    });

    return {
      properties: {
        page: {
          size: {
            width: Math.round(page.widthPt * 20),
            height: Math.round(page.heightPt * 20),
          },
          margin: {
            top: Math.round((margins.top ?? 72) * 20),
            right: Math.round((margins.right ?? 72) * 20),
            bottom: Math.round((margins.bottom ?? 72) * 20),
            left: Math.round((margins.left ?? 72) * 20),
          },
        },
      },
      children,
    };
  });

  const document = new Document({
    sections,
  });

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const buffer = await Packer.toBuffer(document);
  await fs.writeFile(outputPath, buffer);
}
module.exports = {
  generateWordDocument,
  renderTableFromLayout: makeTable,
};

const { readDocumentXml } = require("./docx-layout-reader");

const WORD_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function childElements(element, localName) {
  if (!element) return [];

  return Array.from(element.childNodes).filter(
    (node) =>
      node.nodeType === 1 &&
      node.namespaceURI === WORD_NAMESPACE &&
      (!localName || node.localName === localName),
  );
}

function firstChild(element, localName) {
  return childElements(element, localName)[0] ?? null;
}

function wordAttribute(element, name) {
  return element?.getAttributeNS(WORD_NAMESPACE, name) || null;
}

function borderValue(element) {
  if (!element) return null;

  const style = wordAttribute(element, "val") || "single";

  if (style === "nil" || style === "none") {
    return style;
  }

  const border = { style };
  const color = wordAttribute(element, "color");
  const size = Number(wordAttribute(element, "sz"));

  if (color && /^[0-9a-f]{6}$/i.test(color)) {
    border.color = color;
  }

  if (Number.isFinite(size) && size > 0) {
    border.size = size;
  }

  return border;
}

function readBorders(properties, sideNames, fallbackProperties = []) {
  const result = {};

  for (const source of [...fallbackProperties, properties]) {
    const bordersElement =
      firstChild(source, "tcBorders") || firstChild(source, "tblBorders");

    if (!bordersElement) continue;

    for (const [xmlName, jsonName] of sideNames) {
      const borderElement = firstChild(bordersElement, xmlName);

      if (borderElement) {
        result[jsonName] = borderValue(borderElement);
      }
    }
  }

  return Object.keys(result).length ? result : undefined;
}

function readMargins(marginsElement) {
  if (!marginsElement) return undefined;

  const margins = {};

  for (const side of ["top", "right", "bottom", "left"]) {
    const value = Number(wordAttribute(firstChild(marginsElement, side), "w"));
    if (Number.isFinite(value)) margins[side] = value;
  }

  return Object.keys(margins).length ? margins : undefined;
}

function paragraphAlignment(paragraph) {
  const paragraphProperties = firstChild(paragraph, "pPr");
  const alignment = wordAttribute(firstChild(paragraphProperties, "jc"), "val");

  if (["left", "center", "right"].includes(alignment)) {
    return alignment;
  }

  return undefined;
}

function readSectionSettings(sectionProperties) {
  const pageSize = firstChild(sectionProperties, "pgSz");
  const pageMargins = firstChild(sectionProperties, "pgMar");

  const page = {
    widthPt: Number(wordAttribute(pageSize, "w")) / 20,
    heightPt: Number(wordAttribute(pageSize, "h")) / 20,
  };

  const marginsPt = {};

  for (const side of ["top", "right", "bottom", "left"]) {
    const value = Number(wordAttribute(pageMargins, side));
    if (Number.isFinite(value)) marginsPt[side] = value / 20;
  }

  return {
    page,
    marginsPt,
  };
}

function readCell(cell, rowProperties) {
  const properties = firstChild(cell, "tcPr");
  const span = Number(
    wordAttribute(firstChild(properties, "gridSpan"), "val") || 1,
  );
  const shading = wordAttribute(firstChild(properties, "shd"), "fill");
  const verticalAlignment = wordAttribute(
    firstChild(properties, "vAlign"),
    "val",
  );
  const verticalMerge = firstChild(properties, "vMerge");
  const nestedTable = firstChild(cell, "tbl");
  const firstParagraph = firstChild(cell, "p");

  const output = nestedTable
    ? { ...readTable(nestedTable), type: "table" }
    : { value: "" };

  if (span > 1) output.columnSpan = span;
  if (shading && shading !== "auto") output.shading = shading;

  if (["top", "center", "bottom"].includes(verticalAlignment)) {
    output.verticalAlignment = verticalAlignment;
  }

  if (verticalMerge) {
    output.verticalMerge = wordAttribute(verticalMerge, "val") || "continue";
  }

  const alignment = paragraphAlignment(firstParagraph);
  if (alignment) output.alignment = alignment;

  const rowBorderProperties = firstChild(rowProperties, "tblPrEx");
  const borders = readBorders(
    properties,
    [
      ["top", "top"],
      ["bottom", "bottom"],
      ["left", "left"],
      ["right", "right"],
      ["start", "start"],
      ["end", "end"],
    ],
    [rowBorderProperties],
  );

  if (borders) output.borders = borders;

  return output;
}

function readTable(table) {
  const properties = firstChild(table, "tblPr");
  const grid = firstChild(table, "tblGrid");
  const gridWidths = childElements(grid, "gridCol").map((column) =>
    Number(wordAttribute(column, "w")),
  );
  const totalGridWidth = gridWidths.reduce((sum, width) => sum + width, 0);

  const tableWidth = firstChild(properties, "tblW");
  const widthDxa = Number(wordAttribute(tableWidth, "w"));
  const widthType = wordAttribute(tableWidth, "type");

  const columnWidthsPercent =
    totalGridWidth > 0
      ? gridWidths.map((width) => (width / totalGridWidth) * 100)
      : undefined;

  const rows = childElements(table, "tr").map((row) => {
    const rowProperties = firstChild(row, "trPr");
    const height = Number(
      wordAttribute(firstChild(rowProperties, "trHeight"), "val"),
    );

    const outputRow = {
      cells: childElements(row, "tc").map((cell) =>
        readCell(cell, rowProperties),
      ),
    };

    if (Number.isFinite(height) && height > 0) {
      outputRow.heightTwips = height;
    }

    if (firstChild(rowProperties, "tblHeader")) {
      outputRow.header = true;
    }

    if (firstChild(rowProperties, "cantSplit")) {
      outputRow.cantSplit = true;
    }

    return outputRow;
  });

  const output = { rows };

  if (columnWidthsPercent) output.columnWidthsPercent = columnWidthsPercent;
  if (widthType === "dxa" && Number.isFinite(widthDxa)) {
    output.widthDxa = widthDxa;
  }

  const indent = Number(wordAttribute(firstChild(properties, "tblInd"), "w"));

  if (Number.isFinite(indent) && indent > 0) {
    output.indentTwips = indent;
  }

  const borders = readBorders(properties, [
    ["top", "top"],
    ["bottom", "bottom"],
    ["left", "left"],
    ["right", "right"],
    ["insideH", "insideHorizontal"],
    ["insideV", "insideVertical"],
  ]);

  if (borders) output.borders = borders;

  const margins = readMargins(firstChild(properties, "tblCellMar"));
  if (margins) output.margins = margins;

  return output;
}

function readParagraph(paragraph) {
  const paragraphProperties = firstChild(paragraph, "pPr");
  const output = {
    type: "text",
    role: "blank-paragraph",
    value: "",
  };

  const alignment = paragraphAlignment(paragraph);
  if (alignment) output.alignment = alignment;

  if (firstChild(paragraphProperties, "pageBreakBefore")) {
    output.pageBreakBefore = true;
  }

  return output;
}

async function extractLayout(docxBuffer) {
  const xmlDocument = await readDocumentXml(docxBuffer);
  const body = firstChild(xmlDocument.documentElement, "body");

  if (!body) {
    throw new Error("The Word document body was not found.");
  }

  const sections = [];
  let blocks = [];

  function finishSection(sectionProperties) {
    const settings = readSectionSettings(sectionProperties);

    sections.push({
      id: `section-${sections.length + 1}`,
      marginsPt: settings.marginsPt,
      blocks,
    });

    blocks = [];
    return settings;
  }

  let firstSectionSettings = null;

  for (const element of childElements(body)) {
    if (element.localName === "p") {
      blocks.push(readParagraph(element));

      const paragraphProperties = firstChild(element, "pPr");
      const sectionProperties = firstChild(paragraphProperties, "sectPr");

      if (sectionProperties) {
        const settings = finishSection(sectionProperties);
        firstSectionSettings ??= settings;
      }
    } else if (element.localName === "tbl") {
      blocks.push({
        ...readTable(element),
        type: "table",
      });
    } else if (element.localName === "sectPr") {
      const settings = finishSection(element);
      firstSectionSettings ??= settings;
    }
  }

  if (blocks.length > 0 || sections.length === 0) {
    const settings = finishSection(null);
    firstSectionSettings ??= settings;
  }

  const page = firstSectionSettings?.page ?? {
    widthPt: 612,
    heightPt: 792,
  };

  return {
    schemaVersion: "1.0",
    source: "uploaded.docx",
    description: "Structure-only layout extracted from a Word document.",
    document: {
      page,
      marginsPt: firstSectionSettings?.marginsPt ?? {},
    },
    sections,
  };
}

module.exports = { extractLayout };

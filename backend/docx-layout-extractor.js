const JSZip = require("jszip");
const crypto = require("node:crypto");
const { xml2js } = require("xml-js");

const WORD_PREFIX = "w:";

function children(node, localName) {
  if (!node) return [];
  return (node.elements ?? []).filter(
    (element) =>
      element.type === "element" &&
      element.name === `${WORD_PREFIX}${localName}`,
  );
}

function first(node, localName) {
  return children(node, localName)[0];
}

function attr(node, localName) {
  if (!node || node.attributes == null) return undefined;
  return (
    node.attributes[`${WORD_PREFIX}${localName}`] ?? node.attributes[localName]
  );
}

function numberAttr(node, name) {
  const value = Number(attr(node, name));
  return Number.isFinite(value) ? value : undefined;
}

function borderObject(element) {
  if (!element) return undefined;
  const style = attr(element, "val") ?? "single";
  if (style === "nil" || style === "none") return "nil";

  const border = { style };
  const color = attr(element, "color");
  const size = numberAttr(element, "sz");
  if (color && /^[\da-f]{6}$/i.test(color)) border.color = color;
  if (size !== undefined && size > 0) border.size = size;
  return border;
}

const CELL_EDGES = ["top", "bottom", "left", "right", "start", "end"];
const TABLE_EDGES = {
  top: "top",
  bottom: "bottom",
  left: "left",
  right: "right",
  insideH: "insideHorizontal",
  insideV: "insideVertical",
  start: "start",
  end: "end",
};

function bordersFrom(element, containerName, edges) {
  const container = first(element, containerName);
  if (!container) return {};
  const borders = {};
  for (const edge of edges) {
    const declaration = first(container, edge);
    if (declaration) borders[edge] = borderObject(declaration);
  }
  return borders;
}

function sameBorder(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function effectiveCellOverrides(cellProps, rowExceptions, tableBorders) {
  const rowBorders = bordersFrom(rowExceptions, "tblBorders", CELL_EDGES);
  const cellBorders = bordersFrom(cellProps, "tcBorders", CELL_EDGES);
  const result = {};

  for (const edge of CELL_EDGES) {
    if (Object.hasOwn(cellBorders, edge)) {
      result[edge] = cellBorders[edge];
      continue;
    }
    if (Object.hasOwn(rowBorders, edge)) {
      const tableEdge = tableBorders[edge];
      if (
        !Object.hasOwn(tableBorders, edge) ||
        !sameBorder(rowBorders[edge], tableEdge)
      ) {
        result[edge] = rowBorders[edge];
      }
    }
  }
  return Object.keys(result).length ? result : undefined;
}

function paragraphAlignment(paragraph) {
  const alignment = attr(first(first(paragraph, "pPr"), "jc"), "val");
  return ["left", "center", "right"].includes(alignment)
    ? alignment
    : undefined;
}

function cellShading(properties) {
  const shading = first(properties, "shd");
  const fill = attr(shading, "fill");
  return fill && /^[\da-f]{6}$/i.test(fill) ? fill : undefined;
}

function cellSpan(properties) {
  return numberAttr(first(properties, "gridSpan"), "val") ?? 1;
}

function makeCell(cell, rowExceptions, tableBorders) {
  const properties = first(cell, "tcPr");
  const paragraph = first(cell, "p");
  const result = { value: "" };
  const span = cellSpan(properties);
  if (span > 1) result.columnSpan = span;
  const width = numberAttr(first(properties, "tcW"), "w");
  if (width !== undefined && width > 0) result.widthDxa = width;
  const shading = cellShading(properties);
  if (shading) result.shading = shading;
  const alignment = paragraphAlignment(paragraph);
  if (alignment) result.alignment = alignment;
  const borders = effectiveCellOverrides(
    properties,
    rowExceptions,
    tableBorders,
  );
  if (borders) result.borders = borders;
  return result;
}

function parseTable(table) {
  const properties = first(table, "tblPr");
  const tableBorders = bordersFrom(properties, "tblBorders", CELL_EDGES);
  const rowNodes = children(table, "tr");
  const grid = first(table, "tblGrid");
  const gridWidths = children(grid, "gridCol").map(
    (column) => numberAttr(column, "w") ?? 0,
  );
  const totalGridWidth = gridWidths.reduce((sum, width) => sum + width, 0);
  const declaredWidth = numberAttr(first(properties, "tblW"), "w");
  const widthDxa = declaredWidth > 0 ? declaredWidth : totalGridWidth;
  const borderValues = {};

  for (const [xmlEdge, jsonEdge] of Object.entries(TABLE_EDGES)) {
    const border = bordersFrom(properties, "tblBorders", [xmlEdge])[xmlEdge];
    if (border !== undefined) borderValues[jsonEdge] = border;
  }

  const rows = rowNodes.map((row) => {
    const rowProperties = first(row, "trPr");
    const rowExceptions = first(row, "tblPrEx");
    const height = numberAttr(first(rowProperties, "trHeight"), "val");
    return {
      ...(height ? { heightTwips: height } : {}),
      ...(first(rowProperties, "cantSplit") ? { cantSplit: true } : {}),
      cells: children(row, "tc").map((cell) =>
        makeCell(cell, rowExceptions, tableBorders),
      ),
    };
  });

  const columnWidthsPercent =
    gridWidths.length && totalGridWidth
      ? gridWidths.map((width) => (width / totalGridWidth) * 100)
      : undefined;

  return {
    type: "table",
    ...(widthDxa ? { widthDxa } : {}),
    ...(numberAttr(first(properties, "tblInd"), "w")
      ? { indentTwips: numberAttr(first(properties, "tblInd"), "w") }
      : {}),
    ...(columnWidthsPercent ? { columnWidthsPercent } : {}),
    ...(Object.keys(borderValues).length ? { borders: borderValues } : {}),
    rows,
  };
}

function sectionPage(documentXml) {
  const body = first(documentXml, "body");
  const section = first(body, "sectPr");
  const size = first(section, "pgSz");
  const margins = first(section, "pgMar");
  const width = numberAttr(size, "w");
  const height = numberAttr(size, "h");
  const page = {
    widthPt: width ? width / 20 : 595.28,
    heightPt: height ? height / 20 : 841.89,
  };
  const marginsPt = {};
  for (const edge of ["top", "right", "bottom", "left"]) {
    const value = numberAttr(margins, edge);
    if (value !== undefined) marginsPt[edge] = value / 20;
  }
  return { page, marginsPt };
}

function descendants(node, localName) {
  const result = [];
  const expectedName = `${WORD_PREFIX}${localName}`;
  const visit = (current) => {
    if (!current || current.type !== "element") return;
    if (current.name === expectedName) result.push(current);
    for (const child of current.elements ?? []) visit(child);
  };
  visit(node);
  return result;
}

function sectionDescriptor(sectionProperties, id, blocks) {
  const size = first(sectionProperties, "pgSz");
  const margins = first(sectionProperties, "pgMar");
  const widthTwips = numberAttr(size, "w");
  const heightTwips = numberAttr(size, "h");
  const orientation =
    attr(size, "orient") ??
    (widthTwips > heightTwips ? "landscape" : "portrait");
  const marginsPt = {};
  for (const edge of ["top", "right", "bottom", "left"]) {
    const value = numberAttr(margins, edge);
    if (value !== undefined) marginsPt[edge] = value / 20;
  }

  return {
    id: `section-${id}`,
    startType: attr(first(sectionProperties, "type"), "val") ?? "nextPage",
    orientation,
    page: {
      widthPt: widthTwips ? widthTwips / 20 : 612,
      heightPt: heightTwips ? heightTwips / 20 : 792,
    },
    marginsPt,
    blocks,
  };
}

function extractSections(documentNode, body) {
  const sections = [];
  let currentBlocks = [];
  let blockIndex = 0;
  let tableIndex = 0;

  const finishSection = (sectionProperties) => {
    sections.push(
      sectionDescriptor(sectionProperties, sections.length + 1, currentBlocks),
    );
    currentBlocks = [];
  };

  for (const element of body.elements ?? []) {
    if (element.type !== "element") continue;
    if (element.name === `${WORD_PREFIX}p`) {
      currentBlocks.push({ type: "paragraph", index: blockIndex++ });
      const sectionProperties = first(first(element, "pPr"), "sectPr");
      if (sectionProperties) finishSection(sectionProperties);
    } else if (element.name === `${WORD_PREFIX}tbl`) {
      currentBlocks.push({ type: "table", tableIndex, index: blockIndex++ });
      tableIndex += 1;
    } else if (element.name === `${WORD_PREFIX}sectPr`) {
      finishSection(element);
    }
  }

  if (currentBlocks.length) {
    const lastSectionProperties = descendants(documentNode, "sectPr").at(-1);
    finishSection(lastSectionProperties);
  }

  return sections;
}

function tableOutline(table, index) {
  const parsed = parseTable(table);
  const allTables = descendants(table, "tbl").length;
  return {
    index,
    rowCount: parsed.rows.length,
    cellCount: parsed.rows.reduce((sum, row) => sum + row.cells.length, 0),
    nestedTableCount: Math.max(0, allTables - 1),
    widthDxa: parsed.widthDxa,
    indentTwips: parsed.indentTwips,
    columnWidthsPercent: parsed.columnWidthsPercent,
    borders: parsed.borders,
    rows: parsed.rows,
  };
}

async function extractDocxLayout(buffer, sourceName = "uploaded.docx") {
  const zip = await JSZip.loadAsync(buffer);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile)
    throw new Error("The uploaded file is not a valid Word document.");

  const documentXml = xml2js(await documentFile.async("string"), {
    compact: false,
    spaces: 0,
  });
  const documentNode = first(documentXml, "document");
  const body = first(documentNode, "body");
  if (!body) throw new Error("The Word document has no document body.");
  const sections = extractSections(documentNode, body);
  const allTables = descendants(body, "tbl");
  const topLevelTables = children(body, "tbl");
  const paragraphs = descendants(body, "p");
  const tableRows = descendants(body, "tr");
  const tableCells = descendants(body, "tc");
  const drawings = descendants(body, "drawing");
  const pictures = descendants(body, "pict");
  const imageParts = Object.keys(zip.files).filter(
    (name) => name.startsWith("word/media/") && !zip.files[name].dir,
  );
  const allBorderContainers = [
    ...descendants(body, "tcBorders"),
    ...descendants(body, "tblBorders"),
  ];
  const hiddenBorderCount = allBorderContainers.reduce(
    (count, container) =>
      count +
      (container.elements ?? []).filter(
        (edge) =>
          edge.type === "element" &&
          ["nil", "none"].includes(attr(edge, "val")),
      ).length,
    0,
  );
  const firstSection = sections[0] ?? sectionDescriptor(undefined, 1, []);
  const packageParts = await Promise.all(
    Object.entries(zip.files)
      .filter(([, file]) => !file.dir)
      .map(async ([name, file]) => {
        const bytes = await file.async("nodebuffer");
        const isXml =
          /\.(?:xml|rels)$/i.test(name) || name === "[Content_Types].xml";
        return {
          path: name,
          encoding: isXml ? "utf8" : "base64",
          content: isXml ? bytes.toString("utf8") : bytes.toString("base64"),
          sha256: crypto.createHash("sha256").update(bytes).digest("hex"),
          date: file.date?.toISOString(),
        };
      }),
  );

  return {
    schemaVersion: "2.0",
    source: sourceName,
    description:
      "Complete Word package snapshot with a structure index; source content and formatting are preserved.",
    document: { page: firstSection.page, marginsPt: firstSection.marginsPt },
    sections,
    structure: {
      sectionCount: sections.length,
      topLevelParagraphCount: children(body, "p").length,
      nestedParagraphCount: paragraphs.length - children(body, "p").length,
      topLevelTableCount: topLevelTables.length,
      tableCountIncludingNested: allTables.length,
      nestedTableCount: allTables.length - topLevelTables.length,
      rowCount: tableRows.length,
      cellCount: tableCells.length,
      drawingCount: drawings.length,
      pictureCount: pictures.length,
      mediaPartCount: imageParts.length,
      explicitHiddenBorderCount: hiddenBorderCount,
      tables: topLevelTables.map(tableOutline),
    },
    packagePayload: {
      encoding: "docx-parts-json",
      sourceSha256: crypto.createHash("sha256").update(buffer).digest("hex"),
      parts: packageParts,
    },
  };
}

module.exports = {
  extractLayout: extractDocxLayout,
  extractDocxLayout,
};

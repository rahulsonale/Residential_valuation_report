const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");

const WORD_NAMESPACE =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

async function readDocumentXml(docxBuffer) {
  if (!Buffer.isBuffer(docxBuffer)) {
    throw new TypeError("Expected the uploaded Word file as a buffer.");
  }

  let zip;

  try {
    zip = await JSZip.loadAsync(docxBuffer);
  } catch {
    throw new Error("The uploaded file is not a readable .docx package.");
  }

  const documentFile = zip.file("word/document.xml");

  if (!documentFile) {
    throw new Error(
      "The Word document XML was not found in the .docx package.",
    );
  }

  const xmlText = await documentFile.async("string");

  let xmlDocument;

  try {
    xmlDocument = new DOMParser().parseFromString(xmlText, "application/xml");
  } catch {
    throw new Error("The Word document XML could not be read.");
  }

  const root = xmlDocument.documentElement;

  if (
    !root ||
    root.localName !== "document" ||
    root.namespaceURI !== WORD_NAMESPACE
  ) {
    throw new Error(
      "The uploaded file does not contain a valid Word document.",
    );
  }

  return xmlDocument;
}

module.exports = { readDocumentXml };

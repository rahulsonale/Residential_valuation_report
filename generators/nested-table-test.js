const fs = require("node:fs/promises");
const path = require("node:path");
const { Document, Packer, Paragraph, TextRun } = require("docx");
const { renderTableFromLayout } = require("./word-generator");

async function main() {
  const projectDir = path.join(__dirname, "..");
  const layoutPath = path.join(projectDir, "config", "report-layout-v2.json");
  const outputPath = path.join(
    projectDir,
    "output",
    "nested-table-demo-v2.docx",
  );

  const layout = JSON.parse(await fs.readFile(layoutPath, "utf8"));
  const reportSection = layout.sections.find(
    (section) => section.id === "report-details",
  );
  const propertyTable = reportSection.blocks.find(
    (block) => block.id === "property-details-table",
  );

  const document = new Document({
    sections: [
      {
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: "Nested Table Support Demo",
                bold: true,
              }),
            ],
          }),
          renderTableFromLayout(propertyTable),
        ],
      },
    ],
  });

  const buffer = await Packer.toBuffer(document);
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await fs.writeFile(outputPath, buffer);

  console.log(`Created ${outputPath}`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

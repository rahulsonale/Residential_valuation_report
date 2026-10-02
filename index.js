const fs = require("node:fs");
const path = require("node:path");
const { generateWordDocument } = require("./generators/word-generator");

const configDir = path.join(__dirname, "config");
const outputPath = path.join(
  __dirname,
  "output",
  "residential-valuation-report-v2.docx",
);

const reportData = JSON.parse(
  fs.readFileSync(path.join(configDir, "report-data.json"), "utf8"),
);

const reportLayout = JSON.parse(
  fs.readFileSync(path.join(configDir, "report-layout-v2.json"), "utf8"),
);

async function main() {
  await generateWordDocument({ reportData, reportLayout, outputPath });
  console.log(`Report created at: ${outputPath}`);
}

main().catch((error) => {
  console.error("Could not generate the report:", error);
  process.exitCode = 1;
});

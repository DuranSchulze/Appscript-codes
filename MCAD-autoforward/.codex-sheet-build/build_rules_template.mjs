import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const projectDir = "/Users/zafajardo/Documents/Development/Appscript-codes/MCAD-autoforward";
const outputDir = path.join(
  projectDir,
  "outputs/019f6d9d-2cb4-75c3-8a38-6982dd675716"
);

function extractRules(source) {
  const startMarker = "RULES: [";
  const start = source.indexOf(startMarker);
  const end = source.indexOf("\n]\n};", start);

  if (start < 0 || end < 0) {
    throw new Error("Could not locate the RULES array.");
  }

  const arrayText = source.slice(
    start + "RULES: ".length,
    end + 2
  );

  return Function(`"use strict"; return (${arrayText});`)();
}

function rulesToRows(rules) {
  return rules.map((rule) => [
    "TRUE",
    rule.sender,
    rule.matchAll ? "TRUE" : "FALSE",
    (rule.keywords || []).join("\n"),
    (rule.recipients || []).join("\n"),
    ""
  ]);
}

function styleRulesSheet(sheet, rows, tableName) {
  const headers = [[
    "Enabled",
    "Sender",
    "Match All",
    "Keywords (one per line)",
    "Recipients (one per line)",
    "Notes"
  ]];

  sheet.getRange("A1:F1").values = headers;
  sheet.getRange(`A2:F${rows.length + 1}`).values = rows;
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);

  const fullRange = sheet.getRange(`A1:F${rows.length + 1}`);
  fullRange.format = {
    font: { name: "Arial", size: 10, color: "#1F2937" },
    verticalAlignment: "top"
  };

  const header = sheet.getRange("A1:F1");
  header.format = {
    fill: "#0F766E",
    font: { name: "Arial", size: 10, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    rowHeightPx: 34,
    borders: { preset: "outside", style: "medium", color: "#0B5F59" }
  };

  sheet.getRange(`A2:A${rows.length + 1}`).format.horizontalAlignment = "center";
  sheet.getRange(`C2:C${rows.length + 1}`).format.horizontalAlignment = "center";
  sheet.getRange(`D2:F${rows.length + 1}`).format.wrapText = true;

  sheet.getRange(`A2:A${rows.length + 1}`).dataValidation = {
    rule: { type: "list", values: ["TRUE", "FALSE"] }
  };
  sheet.getRange(`C2:C${rows.length + 1}`).dataValidation = {
    rule: { type: "list", values: ["TRUE", "FALSE"] }
  };

  sheet.getRange(`A2:F${rows.length + 1}`).conditionalFormats.addCustom(
    '=$A2="FALSE"',
    {
      fill: "#F3F4F6",
      font: { color: "#9CA3AF", italic: true }
    }
  );
  sheet.getRange(`C2:C${rows.length + 1}`).conditionalFormats.addCustom(
    '=$C2="TRUE"',
    {
      fill: "#DBEAFE",
      font: { color: "#1D4ED8", bold: true }
    }
  );

  sheet.getRange("A:A").format.columnWidthPx = 80;
  sheet.getRange("B:B").format.columnWidthPx = 230;
  sheet.getRange("C:C").format.columnWidthPx = 95;
  sheet.getRange("D:D").format.columnWidthPx = 235;
  sheet.getRange("E:E").format.columnWidthPx = 340;
  sheet.getRange("F:F").format.columnWidthPx = 190;

  rows.forEach((row, index) => {
    const lineCount = Math.max(
      1,
      String(row[3]).split("\n").length,
      String(row[4]).split("\n").length
    );
    const height = Math.min(128, Math.max(32, 18 + lineCount * 14));
    sheet.getRange(`${index + 2}:${index + 2}`).format.rowHeightPx = height;
  });

  const table = sheet.tables.add(
    `A1:F${rows.length + 1}`,
    true,
    tableName
  );
  table.style = "TableStyleMedium2";
  table.showBandedRows = true;
  table.showFilterButton = true;
}

await fs.mkdir(outputDir, { recursive: true });

const [code1, code2] = await Promise.all([
  fs.readFile(path.join(projectDir, "code.gs"), "utf8"),
  fs.readFile(path.join(projectDir, "code2.gs"), "utf8")
]);

const sources = [
  {
    name: "Rules - Code.gs",
    tableName: "CodeGsRules",
    rules: extractRules(code1)
  },
  {
    name: "Rules - Code2.gs",
    tableName: "Code2GsRules",
    rules: extractRules(code2)
  }
];

const workbook = Workbook.create();

for (const source of sources) {
  const sheet = workbook.worksheets.add(source.name);
  styleRulesSheet(sheet, rulesToRows(source.rules), source.tableName);

  const preview = await workbook.render({
    sheetName: source.name,
    range: `A1:F${source.rules.length + 1}`,
    scale: 1.2,
    format: "png"
  });
  const previewName = source.name.replace(/[^a-z0-9]+/gi, "_").toLowerCase();
  await fs.writeFile(
    path.join(outputDir, `${previewName}.png`),
    new Uint8Array(await preview.arrayBuffer())
  );
}

for (const source of sources) {
  const inspection = await workbook.inspect({
    kind: "table",
    range: `'${source.name}'!A1:F${source.rules.length + 1}`,
    include: "values,formulas",
    tableMaxRows: 12,
    tableMaxCols: 6
  });
  console.log(inspection.ndjson);
}

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan"
});
console.log(errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
const outputPath = path.join(outputDir, "AutoForward_Rules_Template.xlsx");
await output.save(outputPath);
console.log(outputPath);

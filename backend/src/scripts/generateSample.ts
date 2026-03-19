import * as XLSX from "xlsx";
import path from "node:path";
import fs from "node:fs";

const workbook = XLSX.utils.book_new();

const inputSheet = XLSX.utils.aoa_to_sheet([
  ["Input A", "Input B", "Total", "Weighted"],
  [10, 5, { f: "A2+B2" }, { f: "C2*1.5" }],
  [7, 8, { f: "A3+B3" }, { f: "C3*1.5" }],
  [12, 9, { f: "A4+B4" }, { f: "C4*1.5" }]
]);

const summarySheet = XLSX.utils.aoa_to_sheet([
  ["Metric", "Value"],
  ["Grand Total", { f: "SUM(Input!C2:C4)" }],
  ["Weighted Total", { f: "SUM(Input!D2:D4)" }],
  ["Variance", { f: "B3-B2" }],
  ["Cross check", { f: "Input!C2+Input!C3+Input!C4" }]
]);

XLSX.utils.book_append_sheet(workbook, inputSheet, "Input");
XLSX.utils.book_append_sheet(workbook, summarySheet, "Summary");

const outDir = path.resolve(process.cwd(), "..", "samples");
if (!fs.existsSync(outDir)) {
  fs.mkdirSync(outDir, { recursive: true });
}

const outPath = path.join(outDir, "sample-workbook.xlsx");
XLSX.writeFile(workbook, outPath);

// eslint-disable-next-line no-console
console.log(`Sample workbook written to ${outPath}`);

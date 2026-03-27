import * as XLSX from "xlsx";
import path from "node:path";
import { CellValue, ParsedWorkbookData } from "../models/pipeline";

export class ExcelParser {
  parse(sourceFilePath: string, targetFilePath?: string): ParsedWorkbookData {
    const workbook = XLSX.readFile(sourceFilePath, { cellFormula: true, cellNF: false, cellText: false });
    const resolvedTargetPath = targetFilePath ?? sourceFilePath;
    const targetWorkbook =
      resolvedTargetPath !== sourceFilePath
        ? XLSX.readFile(resolvedTargetPath, { cellFormula: true, cellNF: false, cellText: false })
        : workbook;
    const values: Record<string, Record<string, CellValue>> = {};
    const formulaCells: ParsedWorkbookData["formulaCells"] = [];

    for (const sheetName of workbook.SheetNames) {
      const ws = workbook.Sheets[sheetName];
      if (!ws) {
        continue;
      }

      values[sheetName] = {};

      for (const key of Object.keys(ws)) {
        if (key.startsWith("!")) {
          continue;
        }

        const cell = ws[key] as XLSX.CellObject;
        const raw = cell.v;
        if (typeof raw === "number" || typeof raw === "string" || typeof raw === "boolean") {
          values[sheetName][key.toUpperCase()] = raw;
        }

        if (typeof cell.f === "string" && cell.f.trim()) {
          formulaCells.push({
            sheet: sheetName,
            cell: key.toUpperCase(),
            formula: `=${cell.f}`
          });
        }
      }
    }

    return {
      sourceFilePath,
      sourceFileName: path.basename(sourceFilePath),
      targetFilePath: resolvedTargetPath,
      sheetNames: workbook.SheetNames,
      targetSheetNames: targetWorkbook.SheetNames,
      formulaCells,
      values
    };
  }
}

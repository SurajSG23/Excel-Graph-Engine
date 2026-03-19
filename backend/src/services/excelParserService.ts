import * as XLSX from "xlsx";
import { ParsedCell } from "../models/graph";

export class ExcelParserService {
  parseWorkbook(filePath: string): { sheets: string[]; cells: ParsedCell[] } {
    const workbook = XLSX.readFile(filePath, { cellFormula: true, cellNF: false, cellText: false });
    const cells: ParsedCell[] = [];

    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        continue;
      }

      for (const key of Object.keys(sheet)) {
        if (key.startsWith("!")) {
          continue;
        }

        const cell = sheet[key] as XLSX.CellObject | undefined;
        if (!cell) {
          continue;
        }

        const value = typeof cell.v === "number" ? cell.v : undefined;
        const formula = cell.f ? `=${cell.f}` : undefined;

        cells.push({
          sheet: sheetName,
          cell: key,
          formula,
          value
        });
      }
    }

    return {
      sheets: workbook.SheetNames,
      cells
    };
  }
}

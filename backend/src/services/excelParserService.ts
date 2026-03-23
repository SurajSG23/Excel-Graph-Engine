import * as XLSX from "xlsx";
import { ParsedWorkbook, WorkbookRole } from "../models/graph";
import { normalizeFileName } from "../utils/cellUtils";

export class ExcelParserService {
  parseWorkbook(filePath: string, uploadName: string, fileRole: WorkbookRole): ParsedWorkbook {
    const workbook = XLSX.readFile(filePath, { cellFormula: true, cellNF: false, cellText: false });
    const fileName = normalizeFileName(uploadName);
    const cells = [] as ParsedWorkbook["cells"];

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

        const raw = cell.v;
        const value =
          typeof raw === "number" || typeof raw === "string" || typeof raw === "boolean"
            ? raw
            : undefined;
        const formula = cell.f ? `=${cell.f}` : undefined;

        cells.push({
          fileName,
          fileRole,
          sheet: sheetName,
          cell: key,
          formula,
          value
        });
      }
    }

    return {
      fileName,
      fileRole,
      uploadName,
      sheets: workbook.SheetNames,
      cells
    };
  }
}

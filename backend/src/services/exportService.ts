import * as XLSX from "xlsx";
import path from "node:path";
import { GraphNode } from "../models/graph";
import { parseRangeRef } from "../utils/cellUtils";

export class ExportService {
  exportWorkbook(nodes: GraphNode[], workbookId: string, outputFileName: string): string {
    const workbook = XLSX.utils.book_new();
    const outputNodes = nodes.filter(
      (node) => node.fileName === outputFileName && (node.nodeType === "output" || node.nodeType === "formula")
    );
    const outputSheets = [...new Set(outputNodes.map((node) => node.sheet))];

    for (const sheetName of outputSheets) {
      const sheetNodes = outputNodes.filter((node) => node.sheet === sheetName);
      const ws: XLSX.WorkSheet = {};

      for (const node of sheetNodes) {
        const parsed = parseRangeRef(node.range);
        if (!parsed) {
          continue;
        }

        const values = node.rangeValues ?? (node.value !== undefined ? [node.value] : []);
        const formulas = node.formulaByCell ?? {};

        for (let index = 0; index < parsed.cells.length; index += 1) {
          const cell = parsed.cells[index];
          const value = values[index] ?? values[0] ?? "";
          const formula = formulas[cell];
          ws[cell] = this.toCellObject(value, formula);
        }
      }

      const addresses = sheetNodes.flatMap((node) => {
        const parsed = parseRangeRef(node.range);
        if (!parsed) {
          return [] as XLSX.CellAddress[];
        }
        return parsed.cells.map((cell) => XLSX.utils.decode_cell(cell));
      });
      if (addresses.length > 0) {
        const maxCol = Math.max(...addresses.map((a) => a.c));
        const maxRow = Math.max(...addresses.map((a) => a.r));
        ws["!ref"] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: maxCol, r: maxRow } });
      }

      XLSX.utils.book_append_sheet(workbook, ws, sheetName);
    }

    const outPath = path.resolve(process.cwd(), "exports", `${workbookId}-export.xlsx`);
    XLSX.writeFile(workbook, outPath);
    return outPath;
  }

  private toCellObject(value: GraphNode["value"] | "", formula?: string): XLSX.CellObject {
    const formulaBody = formula?.startsWith("=") ? formula.slice(1) : formula;

    const cellType: XLSX.CellObject["t"] =
      typeof value === "number" ? "n" : typeof value === "boolean" ? "b" : "s";

    if (formulaBody) {
      return {
        t: cellType,
        f: formulaBody,
        v: value ?? ""
      };
    }

    return {
      t: cellType,
      v: value ?? ""
    };
  }
}

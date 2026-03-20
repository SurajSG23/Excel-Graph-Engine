import * as XLSX from "xlsx";
import path from "node:path";
import { GraphNode } from "../models/graph";

export class ExportService {
  exportWorkbook(nodes: GraphNode[], sheets: string[], workbookId: string): string {
    const workbook = XLSX.utils.book_new();

    for (const sheetName of sheets) {
      const sheetNodes = nodes.filter((node) => node.sheet === sheetName);
      const ws: XLSX.WorkSheet = {};

      for (const node of sheetNodes) {
        ws[node.cell] = this.toCellObject(node);
      }

      const addresses = sheetNodes.map((node) => XLSX.utils.decode_cell(node.cell));
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

  private toCellObject(node: GraphNode): XLSX.CellObject {
    const value = node.value;
    const formulaBody = node.formula?.startsWith("=") ? node.formula.slice(1) : node.formula;

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

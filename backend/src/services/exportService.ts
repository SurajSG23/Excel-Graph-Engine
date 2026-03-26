import * as XLSX from "xlsx";
import path from "node:path";
import { ParsedWorkbookData, PipelineWorkbook } from "../models/pipeline";
import { expandRange } from "../core/node_models";

export class ExportService {
  exportWorkbook(workbook: PipelineWorkbook, parsed: ParsedWorkbookData): string {
    const wb = XLSX.readFile(parsed.targetFilePath, { cellFormula: true });

    for (const node of workbook.config.formulas) {
      const ws = wb.Sheets[node.output.sheet] ?? {};
      wb.Sheets[node.output.sheet] = ws;

      const outputCells = node.outputCells.length > 0 ? node.outputCells : expandRange(node.output.range);
      const results = workbook.nodeResults[node.id] ?? [];
      outputCells.forEach((cell, index) => {
        ws[cell] = this.toCellObject(results[index]);
      });
    }

    const exportPath = path.resolve(process.cwd(), "exports", `pipeline-${workbook.workbookId}.xlsx`);
    XLSX.writeFile(wb, exportPath);
    return exportPath;
  }

  private toCellObject(value: string | number | boolean | undefined): XLSX.CellObject {
    if (typeof value === "number") {
      return { t: "n", v: value, w: String(value) };
    }
    if (typeof value === "boolean") {
      return { t: "b", v: value, w: value ? "TRUE" : "FALSE" };
    }
    const text = value === undefined ? "" : String(value);
    return { t: "s", v: text, w: text };
  }
}

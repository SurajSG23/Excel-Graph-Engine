import * as formulajs from "formulajs";
import { CellValue, FormulaNodeConfig, ParsedWorkbookData } from "../models/pipeline";
import { expandRange, parseCell, parseRefToken, toCell } from "./node_models";

type SheetValues = Record<string, CellValue>;
type ValueMap = Record<string, SheetValues>;

interface ExecuteResult {
  values: ValueMap;
  nodeResults: Record<string, CellValue[]>;
  issues: Array<{ nodeId: string; message: string }>;
}

function cloneValues(values: ValueMap): ValueMap {
  const out: ValueMap = {};
  for (const [sheet, entries] of Object.entries(values)) {
    out[sheet] = { ...entries };
  }
  return out;
}

function stringifyValue(value: unknown): string {
  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if (typeof value === "string") {
    return JSON.stringify(value);
  }
  return "0";
}

function translateFormula(formula: string, anchorCell: string, outputCell: string, outputSheet: string): string {
  const anchor = parseCell(anchorCell);
  const current = parseCell(outputCell);
  if (!anchor || !current) {
    return formula;
  }

  const dRow = current.row - anchor.row;
  const dCol = current.col - anchor.col;

  return formula.replace(
    /((?:'[^']+'|[A-Za-z0-9_\.]+)!|)(\$?)([A-Z]{1,3})(\$?)([0-9]+)/g,
    (_match, sheetPrefix: string, absCol: string, colText: string, absRow: string, rowText: string) => {
      const base = parseCell(`${colText}${rowText}`);
      if (!base) {
        return _match;
      }

      const nextCol = absCol ? base.col : base.col + dCol;
      const nextRow = absRow ? base.row : base.row + dRow;
      if (nextCol < 1 || nextRow < 1) {
        return _match;
      }

      const nextCell = `${toCell(nextRow, nextCol)}`;
      const prefix = sheetPrefix || `${outputSheet}!`;
      return `${prefix}${nextCell}`;
    }
  );
}

function readValue(values: ValueMap, sheet: string, cell: string): CellValue | undefined {
  return values[sheet]?.[cell.toUpperCase()];
}

function evaluateFormula(formula: string, currentSheet: string, values: ValueMap): CellValue {
  let expression = formula.startsWith("=") ? formula.slice(1) : formula;

  expression = expression.replace(
    /((?:'[^']+'|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+):((?:'[^']+'|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g,
    (full) => {
      const [left, right] = full.split(":") as [string, string];
      const leftRef = parseRefToken(left, currentSheet);
      const rightRef = parseRefToken(right, currentSheet);
      if (leftRef.sheet !== rightRef.sheet) {
        const leftValue = readValue(values, leftRef.sheet, leftRef.cell) ?? 0;
        const rightValue = readValue(values, rightRef.sheet, rightRef.cell) ?? 0;
        return `[${stringifyValue(leftValue)},${stringifyValue(rightValue)}]`;
      }

      const cells = expandRange(`${leftRef.cell}:${rightRef.cell}`);
      const refs = cells.map((cell) => stringifyValue(readValue(values, leftRef.sheet, cell) ?? 0));
      return `[${refs.join(",")}]`;
    }
  );

  expression = expression.replace(
    /((?:'[^']+'|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g,
    (token) => {
      const parsed = parseRefToken(token, currentSheet);
      const value = readValue(values, parsed.sheet, parsed.cell);
      return stringifyValue(value ?? 0);
    }
  );

  expression = expression.replace(/\^/g, "**");
  expression = expression.replace(/#(REF!|N\/A|DIV\/0!|VALUE!|NAME\?|NUM!|NULL!)/gi, "0");
  expression = expression.replace(/\b([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g, (_m, name: string) => {
    const upper = name.toUpperCase();
    if (upper === "TRUE") {
      return "true(";
    }
    if (upper === "FALSE") {
      return "false(";
    }
    return `helpers.fn(\"${upper}\")(`;
  });

  const helpers = {
    fn: (name: string) => {
      const fn = (formulajs as Record<string, (...args: unknown[]) => unknown>)[name];
      if (!fn) {
        throw new Error(`Unsupported function: ${name}`);
      }
      return (...args: unknown[]) => fn(...args.flatMap((arg) => (Array.isArray(arg) ? arg : [arg])));
    }
  };

  // eslint-disable-next-line no-new-func
  const run = new Function("helpers", `return (${expression});`) as (h: typeof helpers) => unknown;
  const raw = run(helpers);
  if (typeof raw === "number" || typeof raw === "string" || typeof raw === "boolean") {
    return raw;
  }
  return 0;
}

export class ExecutionEngine {
  execute(parsed: ParsedWorkbookData, formulasInOrder: FormulaNodeConfig[], configValues?: ValueMap): ExecuteResult {
    const values = cloneValues(configValues ?? parsed.values);
    const nodeResults: Record<string, CellValue[]> = {};
    const issues: Array<{ nodeId: string; message: string }> = [];

    for (const node of formulasInOrder) {
      const outputCells = node.outputCells.length > 0 ? node.outputCells : expandRange(node.output.range);
      const computed: CellValue[] = [];

      for (const outCell of outputCells) {
        const translated = translateFormula(node.formula, node.anchorCell, outCell, node.output.sheet);
        let value: CellValue = 0;
        try {
          value = evaluateFormula(translated, node.output.sheet, values);
        } catch (error) {
          const detail = error instanceof Error ? error.message : "Unknown formula evaluation error";
          issues.push({
            nodeId: node.id,
            message: `${node.name} failed at ${node.output.sheet}!${outCell}: ${detail}`
          });
          value = 0;
        }
        if (!values[node.output.sheet]) {
          values[node.output.sheet] = {};
        }
        values[node.output.sheet][outCell] = value;
        computed.push(value);
      }

      nodeResults[node.id] = computed;
    }

    return { values, nodeResults, issues };
  }
}

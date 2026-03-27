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

function splitFunctionArgs(argsText: string): string[] {
  const args: string[] = [];
  let current = "";
  let depth = 0;
  let inString = false;

  for (let i = 0; i < argsText.length; i += 1) {
    const ch = argsText[i];
    if (ch === '"') {
      inString = !inString;
      current += ch;
      continue;
    }

    if (!inString) {
      if (ch === "(") {
        depth += 1;
      } else if (ch === ")") {
        depth = Math.max(0, depth - 1);
      } else if (ch === "," && depth === 0) {
        args.push(current.trim());
        current = "";
        continue;
      }
    }

    current += ch;
  }

  if (current.trim()) {
    args.push(current.trim());
  }

  return args;
}

function resolveScalarToken(token: string, currentSheet: string, values: ValueMap): CellValue {
  const text = token.trim();
  if (!text) {
    return 0;
  }

  if (text.startsWith('"') && text.endsWith('"') && text.length >= 2) {
    return text.slice(1, -1);
  }

  if (/^-?\d+(?:\.\d+)?$/.test(text)) {
    return Number(text);
  }

  if (/^TRUE$/i.test(text)) {
    return true;
  }

  if (/^FALSE$/i.test(text)) {
    return false;
  }

  const refPattern = /^(?:'[^']+'|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+$|^\$?[A-Z]{1,3}\$?[0-9]+$/;
  if (refPattern.test(text)) {
    const parsed = parseRefToken(text, currentSheet);
    return readValue(values, parsed.sheet, parsed.cell) ?? 0;
  }

  return text;
}

function resolveTableRange(token: string, currentSheet: string, values: ValueMap): CellValue[][] {
  const rangeText = token.trim();
  const [left, right] = rangeText.includes(":")
    ? (rangeText.split(":") as [string, string])
    : ([rangeText, rangeText] as [string, string]);

  const leftRef = parseRefToken(left, currentSheet);
  const rightRef = parseRefToken(right, currentSheet);
  const from = parseCell(leftRef.cell);
  const to = parseCell(rightRef.cell);
  if (!from || !to) {
    return [];
  }

  const sheet = leftRef.sheet;
  const minRow = Math.min(from.row, to.row);
  const maxRow = Math.max(from.row, to.row);
  const minCol = Math.min(from.col, to.col);
  const maxCol = Math.max(from.col, to.col);

  const out: CellValue[][] = [];
  for (let row = minRow; row <= maxRow; row += 1) {
    const rowValues: CellValue[] = [];
    for (let col = minCol; col <= maxCol; col += 1) {
      const cell = toCell(row, col);
      rowValues.push(readValue(values, sheet, cell) ?? 0);
    }
    out.push(rowValues);
  }

  return out;
}

function valuesEqual(left: CellValue, right: CellValue): boolean {
  if (typeof left === "number" && typeof right === "number") {
    return left === right;
  }
  return String(left).trim().toLowerCase() === String(right).trim().toLowerCase();
}

function runVlookup(
  lookupValue: CellValue,
  table: CellValue[][],
  colIndex: number,
  approximateMatch: boolean
): CellValue {
  if (table.length === 0) {
    return 0;
  }

  const targetIndex = colIndex - 1;
  if (targetIndex < 0) {
    return 0;
  }

  if (!approximateMatch) {
    for (const row of table) {
      if (row.length === 0) {
        continue;
      }
      if (valuesEqual(row[0], lookupValue)) {
        return row[targetIndex] ?? 0;
      }
    }
    return 0;
  }

  let candidate: CellValue[] | null = null;
  for (const row of table) {
    if (row.length === 0) {
      continue;
    }
    if (valuesEqual(row[0], lookupValue)) {
      return row[targetIndex] ?? 0;
    }

    if (typeof row[0] === "number" && typeof lookupValue === "number" && row[0] <= lookupValue) {
      candidate = row;
    }
  }

  return candidate ? candidate[targetIndex] ?? 0 : 0;
}

function preEvaluateVlookup(expression: string, currentSheet: string, values: ValueMap): string {
  const upper = expression.toUpperCase();
  let cursor = 0;
  let output = "";

  while (cursor < expression.length) {
    const found = upper.indexOf("VLOOKUP(", cursor);
    if (found < 0) {
      output += expression.slice(cursor);
      break;
    }

    output += expression.slice(cursor, found);
    const openParen = found + "VLOOKUP".length;
    let depth = 0;
    let closeParen = -1;
    for (let i = openParen; i < expression.length; i += 1) {
      const ch = expression[i];
      if (ch === "(") {
        depth += 1;
      } else if (ch === ")") {
        depth -= 1;
        if (depth === 0) {
          closeParen = i;
          break;
        }
      }
    }

    if (closeParen < 0) {
      output += expression.slice(found);
      break;
    }

    const argsText = expression.slice(openParen + 1, closeParen);
    const args = splitFunctionArgs(argsText);
    let value: CellValue = 0;
    try {
      if (args.length >= 3) {
        const lookupValue = resolveScalarToken(args[0], currentSheet, values);
        const table = resolveTableRange(args[1], currentSheet, values);
        const colIndexRaw = resolveScalarToken(args[2], currentSheet, values);
        const colIndex = Number(colIndexRaw);
        const rangeLookupRaw = args.length >= 4 ? resolveScalarToken(args[3], currentSheet, values) : true;
        const approximateMatch =
          typeof rangeLookupRaw === "boolean"
            ? rangeLookupRaw
            : !/^FALSE$/i.test(String(rangeLookupRaw));
        value = runVlookup(lookupValue, table, Number.isFinite(colIndex) ? colIndex : 1, approximateMatch);
      }
    } catch {
      value = 0;
    }

    output += stringifyValue(value);
    cursor = closeParen + 1;
  }

  return output;
}

function evaluateFormula(formula: string, currentSheet: string, values: ValueMap): CellValue {
  let expression = formula.startsWith("=") ? formula.slice(1) : formula;

  expression = preEvaluateVlookup(expression, currentSheet, values);

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
      const normalizedName = name.replace(/^_XLFN\./i, "").replace(/^_XLWS\./i, "");
      const fn = (formulajs as Record<string, (...args: unknown[]) => unknown>)[normalizedName];
      if (!fn) {
        throw new Error(`Unsupported function: ${normalizedName}`);
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

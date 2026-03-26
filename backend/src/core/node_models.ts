import { PipelineRange } from "../models/pipeline";

export interface CellAddress {
  row: number;
  col: number;
}

export interface ParsedRef {
  sheet: string;
  cell: string;
}

export function columnToIndex(col: string): number {
  let total = 0;
  for (const ch of col.toUpperCase()) {
    total = total * 26 + (ch.charCodeAt(0) - 64);
  }
  return total;
}

export function indexToColumn(index: number): string {
  let n = index;
  let result = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

export function parseCell(cell: string): CellAddress | null {
  const match = cell.toUpperCase().match(/^\$?([A-Z]{1,3})\$?([0-9]+)$/);
  if (!match) {
    return null;
  }
  return {
    col: columnToIndex(match[1]),
    row: Number(match[2])
  };
}

export function toCell(row: number, col: number): string {
  return `${indexToColumn(col)}${row}`;
}

export function normalizeSheet(sheet: string): string {
  return sheet.trim().replace(/^'|'$/g, "") || "Sheet1";
}

export function parseRefToken(token: string, currentSheet: string): ParsedRef {
  const trimmed = token.trim();
  const [sheetPart, cellPart] = trimmed.includes("!")
    ? (trimmed.split("!") as [string, string])
    : ([currentSheet, trimmed] as [string, string]);

  return {
    sheet: normalizeSheet(sheetPart),
    cell: cellPart.replace(/\$/g, "").toUpperCase()
  };
}

export function extractFormulaRefs(formula: string, currentSheet: string): ParsedRef[] {
  const refs: ParsedRef[] = [];
  const regex = /((?:'[^']+'|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g;

  for (const match of formula.matchAll(regex)) {
    refs.push(parseRefToken(match[0], currentSheet));
  }

  return refs;
}

export function expandRange(range: string): string[] {
  const segments = range
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
  const cells: string[] = [];

  for (const segment of segments) {
    const [left, right] = segment.includes(":")
      ? (segment.split(":") as [string, string])
      : ([segment, segment] as [string, string]);
    const from = parseCell(left);
    const to = parseCell(right);
    if (!from || !to) {
      continue;
    }

    const minRow = Math.min(from.row, to.row);
    const maxRow = Math.max(from.row, to.row);
    const minCol = Math.min(from.col, to.col);
    const maxCol = Math.max(from.col, to.col);

    for (let row = minRow; row <= maxRow; row += 1) {
      for (let col = minCol; col <= maxCol; col += 1) {
        cells.push(toCell(row, col));
      }
    }
  }

  return cells;
}

export function collapseCellsToRange(cells: string[]): string {
  if (cells.length === 0) {
    return "A1:A1";
  }

  const parsed = cells
    .map((cell) => ({ cell: cell.toUpperCase(), parsed: parseCell(cell) }))
    .filter((item): item is { cell: string; parsed: CellAddress } => Boolean(item.parsed));

  if (parsed.length === 0) {
    return "A1:A1";
  }

  const minRow = Math.min(...parsed.map((item) => item.parsed.row));
  const maxRow = Math.max(...parsed.map((item) => item.parsed.row));
  const minCol = Math.min(...parsed.map((item) => item.parsed.col));
  const maxCol = Math.max(...parsed.map((item) => item.parsed.col));

  return `${toCell(minRow, minCol)}:${toCell(maxRow, maxCol)}`;
}

export function rangesToCellKeys(ranges: PipelineRange[]): Set<string> {
  const out = new Set<string>();
  for (const item of ranges) {
    for (const cell of expandRange(item.range)) {
      out.add(`${normalizeSheet(item.sheet)}!${cell}`);
    }
  }
  return out;
}

export function rangesAreValid(ranges: PipelineRange[]): boolean {
  for (const item of ranges) {
    if (!item.sheet || !item.range) {
      return false;
    }
    const expanded = expandRange(item.range);
    if (expanded.length === 0) {
      return false;
    }
  }
  return true;
}

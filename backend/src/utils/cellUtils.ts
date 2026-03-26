const CELL_REF_REGEX = /^\$?([A-Z]{1,3})\$?([0-9]+)$/;
const RANGE_REF_REGEX = /^(\$?[A-Z]{1,3}\$?[0-9]+)(?::(\$?[A-Z]{1,3}\$?[0-9]+))?$/;

export interface ParsedRange {
  startCell: string;
  endCell: string;
  cells: string[];
  rows: number;
  cols: number;
  size: number;
}

export function normalizeSheetName(sheet: string): string {
  return sheet.replace(/^'|'$/g, "").trim();
}

export function normalizeFileName(fileName: string): string {
  const normalized = fileName.trim().replace(/^'|'$/g, "");
  const withoutBrackets = normalized.replace(/^\[/, "").replace(/\]$/, "");
  const parts = withoutBrackets.split(/[\\/]/).filter(Boolean);
  const last = parts.length > 0 ? parts[parts.length - 1] : withoutBrackets;
  return last;
}

export function toNodeId(fileName: string, sheet: string, cell: string): string {
  return `${normalizeFileName(fileName)}::${normalizeSheetName(sheet)}::${normalizeCellAddress(cell)}`;
}

export function splitNodeId(nodeId: string): { fileName: string; sheet: string; cell: string } {
  const [fileName, sheet, cell] = nodeId.split("::");
  return {
    fileName,
    sheet,
    cell
  };
}

export function normalizeCellAddress(cellAddress: string): string {
  return cellAddress.replace(/\$/g, "").toUpperCase();
}

export function colToNumber(col: string): number {
  let result = 0;
  for (let i = 0; i < col.length; i += 1) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result;
}

export function numberToCol(num: number): string {
  let n = num;
  let col = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

export function parseCellRef(ref: string): { col: string; row: number } | null {
  const normalized = normalizeCellAddress(ref);
  const match = normalized.match(CELL_REF_REGEX);
  if (!match) {
    return null;
  }
  return {
    col: match[1],
    row: Number(match[2])
  };
}

export function expandRange(startRef: string, endRef: string): string[] {
  const start = parseCellRef(startRef);
  const end = parseCellRef(endRef);

  if (!start || !end) {
    return [];
  }

  const startCol = colToNumber(start.col);
  const endCol = colToNumber(end.col);
  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);
  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);

  const cells: string[] = [];
  for (let col = minCol; col <= maxCol; col += 1) {
    for (let row = minRow; row <= maxRow; row += 1) {
      cells.push(`${numberToCol(col)}${row}`);
    }
  }

  return cells;
}

export function parseRangeRef(rangeRef: string): ParsedRange | null {
  const normalized = rangeRef.replace(/\s+/g, "").toUpperCase();
  const match = normalized.match(RANGE_REF_REGEX);
  if (!match) {
    return null;
  }

  const startCell = normalizeCellAddress(match[1]);
  const endCell = normalizeCellAddress(match[2] ?? match[1]);
  const start = parseCellRef(startCell);
  const end = parseCellRef(endCell);
  if (!start || !end) {
    return null;
  }

  const startCol = colToNumber(start.col);
  const endCol = colToNumber(end.col);
  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);
  const minRow = Math.min(start.row, end.row);
  const maxRow = Math.max(start.row, end.row);
  const rows = maxRow - minRow + 1;
  const cols = maxCol - minCol + 1;

  return {
    startCell: `${numberToCol(minCol)}${minRow}`,
    endCell: `${numberToCol(maxCol)}${maxRow}`,
    cells: expandRange(`${numberToCol(minCol)}${minRow}`, `${numberToCol(maxCol)}${maxRow}`),
    rows,
    cols,
    size: rows * cols
  };
}

export function encodeRange(startCell: string, endCell?: string): string {
  const start = normalizeCellAddress(startCell);
  const end = normalizeCellAddress(endCell ?? startCell);
  return start === end ? start : `${start}:${end}`;
}

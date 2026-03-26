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

/**
 * Creates a cell-level key for internal value tracking during execution.
 * NOT for graph node IDs - use toRangeNodeId() for graph node identifiers.
 * @internal
 */
export function toCellKey(fileName: string, sheet: string, cell: string): string {
  return `${normalizeFileName(fileName)}::${normalizeSheetName(sheet)}::${normalizeCellAddress(cell)}`;
}

/**
 * @deprecated Use toCellKey() for internal cell tracking or range-based node IDs for graph structure.
 * This function creates cell-level IDs which violates the range-first model.
 */
export function toNodeId(fileName: string, sheet: string, cell: string): string {
  return toCellKey(fileName, sheet, cell);
}

/**
 * Extracts the start cell from a range reference.
 * Used to derive anchor cell from range when needed.
 */
export function getStartCell(range: string): string {
  const parsed = parseRangeRef(range);
  return parsed ? parsed.startCell : normalizeCellAddress(range);
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

/**
 * Checks if two ranges overlap.
 * Used for validation of overlapping writes.
 */
export function rangesOverlap(range1: string, range2: string): boolean {
  const parsed1 = parseRangeRef(range1);
  const parsed2 = parseRangeRef(range2);
  if (!parsed1 || !parsed2) {
    return false;
  }

  const cells1 = new Set(parsed1.cells);
  return parsed2.cells.some((cell) => cells1.has(cell));
}

/**
 * Checks if two ranges are adjacent (share an edge).
 * Used for range merging optimization.
 */
export function rangesAdjacent(range1: string, range2: string): boolean {
  const parsed1 = parseRangeRef(range1);
  const parsed2 = parseRangeRef(range2);
  if (!parsed1 || !parsed2) {
    return false;
  }

  const start1 = parseCellRef(parsed1.startCell);
  const end1 = parseCellRef(parsed1.endCell);
  const start2 = parseCellRef(parsed2.startCell);
  const end2 = parseCellRef(parsed2.endCell);

  if (!start1 || !end1 || !start2 || !end2) {
    return false;
  }

  const col1Start = colToNumber(start1.col);
  const col1End = colToNumber(end1.col);
  const col2Start = colToNumber(start2.col);
  const col2End = colToNumber(end2.col);

  // Check if horizontally adjacent (same rows, columns touch)
  if (start1.row === start2.row && end1.row === end2.row) {
    if (col1End + 1 === col2Start || col2End + 1 === col1Start) {
      return true;
    }
  }

  // Check if vertically adjacent (same columns, rows touch)
  if (col1Start === col2Start && col1End === col2End) {
    if (end1.row + 1 === start2.row || end2.row + 1 === start1.row) {
      return true;
    }
  }

  return false;
}

/**
 * Merges two adjacent ranges into one.
 * Returns null if ranges cannot be merged.
 */
export function mergeRanges(range1: string, range2: string): string | null {
  if (!rangesAdjacent(range1, range2)) {
    return null;
  }

  const parsed1 = parseRangeRef(range1);
  const parsed2 = parseRangeRef(range2);
  if (!parsed1 || !parsed2) {
    return null;
  }

  const start1 = parseCellRef(parsed1.startCell);
  const end1 = parseCellRef(parsed1.endCell);
  const start2 = parseCellRef(parsed2.startCell);
  const end2 = parseCellRef(parsed2.endCell);

  if (!start1 || !end1 || !start2 || !end2) {
    return null;
  }

  const col1Start = colToNumber(start1.col);
  const col1End = colToNumber(end1.col);
  const col2Start = colToNumber(start2.col);
  const col2End = colToNumber(end2.col);

  const minCol = Math.min(col1Start, col2Start);
  const maxCol = Math.max(col1End, col2End);
  const minRow = Math.min(start1.row, start2.row);
  const maxRow = Math.max(end1.row, end2.row);

  return encodeRange(
    `${numberToCol(minCol)}${minRow}`,
    `${numberToCol(maxCol)}${maxRow}`
  );
}

const CELL_REF_REGEX = /^\$?([A-Z]{1,3})\$?([0-9]+)$/;

export function normalizeSheetName(sheet: string): string {
  return sheet.replace(/^'|'$/g, "").trim();
}

export function toNodeId(sheet: string, cell: string): string {
  return `${normalizeSheetName(sheet)}!${normalizeCellAddress(cell)}`;
}

export function splitNodeId(nodeId: string): { sheet: string; cell: string } {
  const [sheet, cell] = nodeId.split("!");
  return {
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

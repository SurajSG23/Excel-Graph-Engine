import {
  GraphNode,
  SpreadsheetMutationResult,
  WorkbookGraph,
  WorkbookOperation,
  WorkbookRole
} from "../models/graph";
import {
  colToNumber,
  normalizeCellAddress,
  normalizeSheetName,
  numberToCol,
  parseCellRef,
  toNodeId
} from "../utils/cellUtils";

const TOKEN_REGEX = /(?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+/g;
const RANGE_REGEX = /((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+):((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g;

interface ParsedToken {
  fileName: string;
  sheet: string;
  cell: string;
}

function parseFormulaToken(token: string, currentFileName: string, currentSheet: string): ParsedToken {
  const cleaned = token.trim();
  if (!cleaned.includes("!")) {
    return {
      fileName: currentFileName,
      sheet: currentSheet,
      cell: normalizeCellAddress(cleaned)
    };
  }

  const [sheetTokenRaw, cellToken] = cleaned.split("!");
  const sheetToken = sheetTokenRaw.replace(/^'|'$/g, "");
  const external = sheetToken.match(/^(?:.*[\\/])?\[([^\]]+)\](.+)$/);

  if (external) {
    return {
      fileName: external[1],
      sheet: normalizeSheetName(external[2]),
      cell: normalizeCellAddress(cellToken)
    };
  }

  return {
    fileName: currentFileName,
    sheet: normalizeSheetName(sheetToken),
    cell: normalizeCellAddress(cellToken)
  };
}

function formatFormulaToken(parsed: ParsedToken, currentFileName: string, currentSheet: string): string {
  if (parsed.fileName === currentFileName && parsed.sheet === currentSheet) {
    return parsed.cell;
  }

  if (parsed.fileName === currentFileName) {
    return `${parsed.sheet}!${parsed.cell}`;
  }

  return `'[${parsed.fileName}]${parsed.sheet}'!${parsed.cell}`;
}

function applyReferenceMapToFormula(
  formula: string | undefined,
  currentFileName: string,
  currentSheet: string,
  nodeIdMap: Map<string, string>
): string | undefined {
  if (!formula || !formula.startsWith("=")) {
    return formula;
  }

  const replaceToken = (token: string): string => {
    const parsed = parseFormulaToken(token, currentFileName, currentSheet);
    const oldId = toNodeId(parsed.fileName, parsed.sheet, parsed.cell);
    const mapped = nodeIdMap.get(oldId);
    if (!mapped) {
      return token;
    }

    const [fileName, sheet, cell] = mapped.split("::");
    return formatFormulaToken({ fileName, sheet, cell }, currentFileName, currentSheet);
  };

  let body = formula.slice(1);
  body = body.replace(RANGE_REGEX, (_match, left, right) => `${replaceToken(left)}:${replaceToken(right)}`);
  body = body.replace(TOKEN_REGEX, (token) => replaceToken(token));

  return `=${body}`;
}

function shiftCellByRow(cell: string, rowDelta: number, thresholdRow: number): string {
  const parsed = parseCellRef(cell);
  if (!parsed || parsed.row < thresholdRow) {
    return cell;
  }

  return `${parsed.col}${Math.max(1, parsed.row + rowDelta)}`;
}

function shiftCellByColumn(cell: string, colDelta: number, thresholdCol: number): string {
  const parsed = parseCellRef(cell);
  if (!parsed) {
    return cell;
  }

  const current = colToNumber(parsed.col);
  if (current < thresholdCol) {
    return cell;
  }

  return `${numberToCol(Math.max(1, current + colDelta))}${parsed.row}`;
}

export class WorkbookMutationService {
  applyOperations(
    workbook: WorkbookGraph,
    operations: WorkbookOperation[]
  ): SpreadsheetMutationResult {
    const workingNodes = workbook.nodes.map((node) => ({ ...node }));
    const files = workbook.files.map((file) => ({ ...file, sheets: [...file.sheets] }));
    const changed = new Set<string>();

    for (const op of operations) {
      switch (op.type) {
        case "ADD_CELL": {
          const id = toNodeId(op.fileName, op.sheet, op.cell);
          const idx = workingNodes.findIndex((node) => node.id === id);
          const role = op.fileRole ?? files.find((file) => file.fileName === op.fileName)?.role ?? "other";
          const created: GraphNode = {
            id,
            fileName: op.fileName,
            fileRole: role,
            sheet: normalizeSheetName(op.sheet),
            cell: normalizeCellAddress(op.cell),
            formula: op.formula,
            value: op.value,
            dependencies: [],
            referenceDetails: []
          };

          if (idx >= 0) {
            workingNodes[idx] = created;
          } else {
            workingNodes.push(created);
          }

          this.ensureSheet(files, op.fileName, op.sheet);
          changed.add(id);
          break;
        }
        case "DELETE_CELLS": {
          const deleteSet = new Set(op.nodeIds);
          for (const id of deleteSet) {
            changed.add(id);
          }
          for (let i = workingNodes.length - 1; i >= 0; i -= 1) {
            if (deleteSet.has(workingNodes[i].id)) {
              workingNodes.splice(i, 1);
            }
          }
          break;
        }
        case "MOVE_CELL": {
          const source = workingNodes.find((node) => node.id === op.fromNodeId);
          if (!source) {
            break;
          }

          const oldId = source.id;
          const normalizedSheet = normalizeSheetName(op.toSheet);
          const targetCell = this.findNextAvailableCell(
            workingNodes,
            op.toFileName,
            normalizedSheet,
            op.toCell,
            oldId
          );
          const newId = toNodeId(op.toFileName, normalizedSheet, targetCell);

          const map = new Map<string, string>([[oldId, newId]]);
          source.fileName = op.toFileName;
          source.sheet = normalizedSheet;
          source.cell = targetCell;
          source.id = newId;
          source.formula = applyReferenceMapToFormula(source.formula, source.fileName, source.sheet, map);
          this.rewriteAllFormulas(workingNodes, map);
          this.ensureSheet(files, op.toFileName, op.toSheet);
          changed.add(oldId);
          changed.add(newId);
          break;
        }
        case "INSERT_ROW": {
          this.shiftRows(workingNodes, files, op.fileName, op.sheet, op.index, op.count ?? 1, changed);
          break;
        }
        case "DELETE_ROW": {
          this.deleteRows(workingNodes, files, op.fileName, op.sheet, op.index, op.count ?? 1, changed);
          break;
        }
        case "INSERT_COLUMN": {
          this.shiftCols(workingNodes, files, op.fileName, op.sheet, op.index, op.count ?? 1, changed);
          break;
        }
        case "DELETE_COLUMN": {
          this.deleteCols(workingNodes, files, op.fileName, op.sheet, op.index, op.count ?? 1, changed);
          break;
        }
        case "ADD_SHEET": {
          this.ensureSheet(files, op.fileName, op.sheet);
          break;
        }
        case "DELETE_SHEET": {
          this.removeSheet(workingNodes, files, op.fileName, op.sheet, changed);
          break;
        }
        case "RENAME_SHEET": {
          this.renameSheet(workingNodes, files, op.fileName, op.fromSheet, op.toSheet, changed);
          break;
        }
        case "COPY_PASTE": {
          this.copyPaste(workingNodes, files, op.sourceNodeIds, op.targetFileName, op.targetSheet, op.targetAnchorCell, changed);
          break;
        }
        default:
          break;
      }
    }

    return {
      nodes: workingNodes,
      files,
      changedNodeIds: [...changed]
    };
  }

  private ensureSheet(files: WorkbookGraph["files"], fileName: string, sheet: string): void {
    const file = files.find((item) => item.fileName === fileName);
    if (!file) {
      files.push({
        fileName,
        role: "other",
        sheets: [sheet],
        uploadName: fileName
      });
      return;
    }

    if (!file.sheets.includes(sheet)) {
      file.sheets.push(sheet);
      file.sheets.sort((a, b) => a.localeCompare(b));
    }
  }

  private rewriteAllFormulas(nodes: GraphNode[], map: Map<string, string>): void {
    for (const node of nodes) {
      node.formula = applyReferenceMapToFormula(node.formula, node.fileName, node.sheet, map);
    }
  }

  private findNextAvailableCell(
    nodes: GraphNode[],
    fileName: string,
    sheet: string,
    preferredCell: string,
    ignoreNodeId?: string
  ): string {
    const normalizedPreferred = normalizeCellAddress(preferredCell);
    const occupied = new Set(
      nodes
        .filter(
          (node) =>
            node.fileName === fileName &&
            node.sheet === sheet &&
            (!ignoreNodeId || node.id !== ignoreNodeId)
        )
        .map((node) => node.cell)
    );

    if (!occupied.has(normalizedPreferred)) {
      return normalizedPreferred;
    }

    const parsed = parseCellRef(normalizedPreferred);
    if (!parsed) {
      let row = 1;
      while (occupied.has(`A${row}`)) {
        row += 1;
      }
      return `A${row}`;
    }

    const startCol = colToNumber(parsed.col);
    let row = parsed.row;

    // Prefer the same column and advance downward to avoid overwriting existing cells.
    while (occupied.has(`${numberToCol(startCol)}${row}`)) {
      row += 1;
    }

    return `${numberToCol(startCol)}${row}`;
  }

  private shiftRows(
    nodes: GraphNode[],
    files: WorkbookGraph["files"],
    fileName: string,
    sheet: string,
    index: number,
    count: number,
    changed: Set<string>
  ): void {
    const map = new Map<string, string>();
    for (const node of nodes) {
      if (node.fileName !== fileName || node.sheet !== sheet) {
        continue;
      }
      const oldId = node.id;
      const nextCell = shiftCellByRow(node.cell, count, index);
      if (nextCell !== node.cell) {
        node.cell = nextCell;
        node.id = toNodeId(node.fileName, node.sheet, node.cell);
        map.set(oldId, node.id);
        changed.add(oldId);
        changed.add(node.id);
      }
    }

    this.ensureSheet(files, fileName, sheet);
    this.rewriteAllFormulas(nodes, map);
  }

  private deleteRows(
    nodes: GraphNode[],
    files: WorkbookGraph["files"],
    fileName: string,
    sheet: string,
    index: number,
    count: number,
    changed: Set<string>
  ): void {
    const maxRow = index + count - 1;
    const map = new Map<string, string>();

    for (let i = nodes.length - 1; i >= 0; i -= 1) {
      const node = nodes[i];
      if (node.fileName !== fileName || node.sheet !== sheet) {
        continue;
      }

      const parsed = parseCellRef(node.cell);
      if (!parsed) {
        continue;
      }

      if (parsed.row >= index && parsed.row <= maxRow) {
        changed.add(node.id);
        nodes.splice(i, 1);
      } else if (parsed.row > maxRow) {
        const oldId = node.id;
        node.cell = `${parsed.col}${parsed.row - count}`;
        node.id = toNodeId(node.fileName, node.sheet, node.cell);
        map.set(oldId, node.id);
        changed.add(oldId);
        changed.add(node.id);
      }
    }

    this.ensureSheet(files, fileName, sheet);
    this.rewriteAllFormulas(nodes, map);
  }

  private shiftCols(
    nodes: GraphNode[],
    files: WorkbookGraph["files"],
    fileName: string,
    sheet: string,
    index: number,
    count: number,
    changed: Set<string>
  ): void {
    const map = new Map<string, string>();

    for (const node of nodes) {
      if (node.fileName !== fileName || node.sheet !== sheet) {
        continue;
      }

      const oldId = node.id;
      const nextCell = shiftCellByColumn(node.cell, count, index);
      if (nextCell !== node.cell) {
        node.cell = nextCell;
        node.id = toNodeId(node.fileName, node.sheet, node.cell);
        map.set(oldId, node.id);
        changed.add(oldId);
        changed.add(node.id);
      }
    }

    this.ensureSheet(files, fileName, sheet);
    this.rewriteAllFormulas(nodes, map);
  }

  private deleteCols(
    nodes: GraphNode[],
    files: WorkbookGraph["files"],
    fileName: string,
    sheet: string,
    index: number,
    count: number,
    changed: Set<string>
  ): void {
    const maxCol = index + count - 1;
    const map = new Map<string, string>();

    for (let i = nodes.length - 1; i >= 0; i -= 1) {
      const node = nodes[i];
      if (node.fileName !== fileName || node.sheet !== sheet) {
        continue;
      }

      const parsed = parseCellRef(node.cell);
      if (!parsed) {
        continue;
      }

      const col = colToNumber(parsed.col);
      if (col >= index && col <= maxCol) {
        changed.add(node.id);
        nodes.splice(i, 1);
      } else if (col > maxCol) {
        const oldId = node.id;
        node.cell = `${numberToCol(col - count)}${parsed.row}`;
        node.id = toNodeId(node.fileName, node.sheet, node.cell);
        map.set(oldId, node.id);
        changed.add(oldId);
        changed.add(node.id);
      }
    }

    this.ensureSheet(files, fileName, sheet);
    this.rewriteAllFormulas(nodes, map);
  }

  private removeSheet(
    nodes: GraphNode[],
    files: WorkbookGraph["files"],
    fileName: string,
    sheet: string,
    changed: Set<string>
  ): void {
    for (let i = nodes.length - 1; i >= 0; i -= 1) {
      if (nodes[i].fileName === fileName && nodes[i].sheet === sheet) {
        changed.add(nodes[i].id);
        nodes.splice(i, 1);
      }
    }

    const file = files.find((entry) => entry.fileName === fileName);
    if (!file) {
      return;
    }

    file.sheets = file.sheets.filter((s) => s !== sheet);
  }

  private renameSheet(
    nodes: GraphNode[],
    files: WorkbookGraph["files"],
    fileName: string,
    fromSheet: string,
    toSheet: string,
    changed: Set<string>
  ): void {
    const map = new Map<string, string>();

    for (const node of nodes) {
      if (node.fileName !== fileName || node.sheet !== fromSheet) {
        continue;
      }

      const oldId = node.id;
      node.sheet = toSheet;
      node.id = toNodeId(node.fileName, node.sheet, node.cell);
      map.set(oldId, node.id);
      changed.add(oldId);
      changed.add(node.id);
    }

    this.rewriteAllFormulas(nodes, map);

    const file = files.find((entry) => entry.fileName === fileName);
    if (!file) {
      return;
    }

    if (!file.sheets.includes(toSheet)) {
      file.sheets.push(toSheet);
    }
    file.sheets = file.sheets.filter((sheet) => sheet !== fromSheet);
    file.sheets.sort((a, b) => a.localeCompare(b));
  }

  private copyPaste(
    nodes: GraphNode[],
    files: WorkbookGraph["files"],
    sourceNodeIds: string[],
    targetFileName: string,
    targetSheet: string,
    targetAnchorCell: string,
    changed: Set<string>
  ): void {
    const sources = nodes.filter((node) => sourceNodeIds.includes(node.id));
    if (sources.length === 0) {
      return;
    }

    const sourceParsed = sources
      .map((node) => ({ node, ref: parseCellRef(node.cell) }))
      .filter((item): item is { node: GraphNode; ref: { col: string; row: number } } => Boolean(item.ref));

    if (sourceParsed.length === 0) {
      return;
    }

    const minRow = Math.min(...sourceParsed.map((item) => item.ref.row));
    const minCol = Math.min(...sourceParsed.map((item) => colToNumber(item.ref.col)));
    const anchor = parseCellRef(targetAnchorCell);
    if (!anchor) {
      return;
    }

    const rowDelta = anchor.row - minRow;
    const colDelta = colToNumber(anchor.col) - minCol;
    const role = files.find((file) => file.fileName === targetFileName)?.role ?? "other";

    for (const item of sourceParsed) {
      const nextRow = Math.max(1, item.ref.row + rowDelta);
      const nextCol = Math.max(1, colToNumber(item.ref.col) + colDelta);
      const cell = `${numberToCol(nextCol)}${nextRow}`;
      const id = toNodeId(targetFileName, targetSheet, cell);

      const pasted: GraphNode = {
        ...item.node,
        fileName: targetFileName,
        fileRole: role,
        sheet: targetSheet,
        cell,
        id
      };

      const existing = nodes.findIndex((node) => node.id === id);
      if (existing >= 0) {
        nodes.splice(existing, 1, pasted);
      } else {
        nodes.push(pasted);
      }
      changed.add(id);
    }

    this.ensureSheet(files, targetFileName, targetSheet);
  }
}

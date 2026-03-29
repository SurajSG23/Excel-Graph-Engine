import { FormulaNodeConfig, ParsedFormulaCell, PipelineRange } from "../models/pipeline";
import { collapseCellsToRange, extractFormulaRefs, normalizeSheet, parseCell, parseRefToken } from "./node_models";

interface GroupBucket {
  sheet: string;
  structureKey: string;
  cells: ParsedFormulaCell[];
}

function keyOf(sheet: string, cell: string): string {
  return `${normalizeSheet(sheet)}!${cell.toUpperCase()}`;
}

function buildStructureKey(formula: string, outputCell: string, outputSheet: string): string {
  const out = parseCell(outputCell);
  if (!out) {
    return formula.trim().toUpperCase();
  }

  return formula
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/((?:'[^']+'|[A-Z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g, (token) => {
      const parsed = parseRefToken(token, outputSheet);
      const pos = parseCell(parsed.cell);
      if (!pos) {
        return token;
      }

      const rawCell = token.includes("!") ? token.split("!")[1] : token;
      const colMatch = rawCell.match(/\$?([A-Z]{1,3})/);
      const rowMatch = rawCell.match(/\$?([0-9]+)/);
      const absCol = rawCell.includes("$") && rawCell.indexOf("$") < rawCell.indexOf(rowMatch?.[0] ?? "");
      const absRow = /[A-Z]{1,3}\$[0-9]+/.test(rawCell);
      const colKey = absCol ? `C${colMatch?.[1] ?? ""}` : `dC${pos.col - out.col}`;
      const rowKey = absRow ? `R${rowMatch?.[1] ?? ""}` : `dR${pos.row - out.row}`;
      return `${normalizeSheet(parsed.sheet)}:${colKey}:${rowKey}`;
    });
}

function splitConnected(cells: ParsedFormulaCell[]): ParsedFormulaCell[][] {
  const map = new Map(cells.map((cell) => [cell.cell, cell]));
  const allCells = [...map.keys()];
  const visited = new Set<string>();
  const groups: ParsedFormulaCell[][] = [];

  for (const cell of cells) {
    if (visited.has(cell.cell)) {
      continue;
    }

    const queue = [cell.cell];
    const chunk: ParsedFormulaCell[] = [];
    visited.add(cell.cell);

    while (queue.length > 0) {
      const current = queue.shift()!;
      const currentPos = parseCell(current);
      const node = map.get(current);
      if (!node || !currentPos) {
        continue;
      }
      chunk.push(node);

      for (const candidate of allCells) {
        if (visited.has(candidate)) {
          continue;
        }
        const pos = parseCell(candidate);
        if (!pos) {
          continue;
        }
        const touching =
          (Math.abs(pos.row - currentPos.row) === 1 && pos.col === currentPos.col) ||
          (Math.abs(pos.col - currentPos.col) === 1 && pos.row === currentPos.row);
        if (touching) {
          visited.add(candidate);
          queue.push(candidate);
        }
      }
    }

    groups.push(chunk);
  }

  return groups;
}

function sortCells(cells: string[]): string[] {
  return [...cells].sort((left, right) => {
    const l = parseCell(left);
    const r = parseCell(right);
    if (!l || !r) {
      return left.localeCompare(right);
    }
    if (l.row !== r.row) {
      return l.row - r.row;
    }
    return l.col - r.col;
  });
}

function buildInputTemplate(mapping: string[]): string {
  if (mapping.length === 0) {
    return "";
  }

  return mapping
    .map((item) => {
      const [sheet, cell] = item.includes("!")
        ? (item.split("!") as [string, string])
        : (["", item] as [string, string]);
      const col = cell.replace(/[0-9]/g, "").replace(/\$/g, "").toUpperCase();
      return sheet ? `${sheet}!${col}` : col;
    })
    .join(", ");
}

export class FormulaGrouper {
  group(formulaCells: ParsedFormulaCell[]): FormulaNodeConfig[] {
    const buckets = new Map<string, GroupBucket>();

    for (const item of formulaCells) {
      const structureKey = buildStructureKey(item.formula, item.cell, item.sheet);
      const bucketKey = keyOf(item.sheet, structureKey);
      if (!buckets.has(bucketKey)) {
        buckets.set(bucketKey, {
          sheet: item.sheet,
          structureKey,
          cells: []
        });
      }
      buckets.get(bucketKey)!.cells.push(item);
    }

    const nodes: FormulaNodeConfig[] = [];
    let index = 1;

    for (const bucket of buckets.values()) {
      for (const connected of splitConnected(bucket.cells)) {
        const outputCells = sortCells(connected.map((cell) => cell.cell));
        const anchorCell = outputCells[0];
        const outputRange = collapseCellsToRange(outputCells);
        const formulaByCell: Record<string, string> = {};
        const inputMappingByCell: Record<string, string[]> = {};
        for (const cell of connected) {
          const outputCell = cell.cell.toUpperCase();
          formulaByCell[outputCell] = cell.formula;
          inputMappingByCell[outputCell] = extractFormulaRefs(cell.formula, cell.sheet).map((ref) => `${ref.sheet}!${ref.cell}`);
        }
        const inputMap = new Map<string, Set<string>>();

        for (const cell of connected) {
          const refs = extractFormulaRefs(cell.formula, cell.sheet);
          for (const ref of refs) {
            if (!inputMap.has(ref.sheet)) {
              inputMap.set(ref.sheet, new Set());
            }
            inputMap.get(ref.sheet)!.add(ref.cell);
          }
        }

        const inputs: PipelineRange[] = [...inputMap.entries()].map(([sheet, refs]) => ({
          sheet,
          range: collapseCellsToRange([...refs])
        }));

        nodes.push({
          id: `formula-${index}`,
          name: `Formula ${index}`,
          inputs,
          inputTemplate: buildInputTemplate(inputMappingByCell[anchorCell.toUpperCase()] ?? []),
          inputMappingByCell,
          output: {
            sheet: bucket.sheet,
            range: outputRange
          },
          formula: connected[0].formula,
          formulaTemplate: connected[0].formula,
          formulaByCell,
          structureKey: bucket.structureKey,
          anchorCell,
          outputCells
        });

        index += 1;
      }
    }

    return nodes;
  }
}

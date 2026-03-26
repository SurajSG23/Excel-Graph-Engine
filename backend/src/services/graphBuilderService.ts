import {
  CellReference,
  GraphEdge,
  GraphNode,
  NodeRangeRef,
  ParsedCell,
  ParsedWorkbook
} from "../models/graph";
import { FormulaParserService } from "./formulaParserService";
import {
  colToNumber,
  encodeRange,
  normalizeCellAddress,
  numberToCol,
  parseCellRef,
  parseRangeRef,
  toNodeId
} from "../utils/cellUtils";

export interface BuildResult {
  nodes: GraphNode[];
  edges: GraphEdge[];
  sheets: string[];
  files: Array<{
    fileName: string;
    role: "input" | "output" | "other";
    sheets: string[];
    uploadName: string;
  }>;
}

export class GraphBuilderService {
  constructor(private readonly formulaParserService: FormulaParserService) {}

  private static readonly FORMULA_TOKEN_REGEX =
    /((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)(:((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+))?/g;

  private resolveExternalFileToken(
    fileToken: string,
    currentFileName: string,
    files: Array<{ fileName: string; role: "input" | "output" | "other" }>
  ): string {
    if (!/^\d+$/.test(fileToken)) {
      return fileToken;
    }

    const rankedCandidates = files
      .filter((file) => file.fileName !== currentFileName)
      .sort((left, right) => {
        const rank = (role: "input" | "output" | "other"): number => {
          if (role === "input") return 0;
          if (role === "output") return 1;
          return 2;
        };

        const byRole = rank(left.role) - rank(right.role);
        if (byRole !== 0) {
          return byRole;
        }
        return left.fileName.localeCompare(right.fileName);
      });

    if (rankedCandidates.length === 0) {
      return fileToken;
    }

    const index = Number(fileToken) - 1;
    if (index >= 0 && index < rankedCandidates.length) {
      return rankedCandidates[index].fileName;
    }

    if (rankedCandidates.length === 1) {
      return rankedCandidates[0].fileName;
    }

    return fileToken;
  }

  buildFromWorkbooks(workbooks: ParsedWorkbook[]): BuildResult {
    const filesMeta = workbooks.map((workbook) => ({ fileName: workbook.fileName, role: workbook.fileRole }));
    const allCells = workbooks.flatMap((workbook) => workbook.cells);
    const formulaCells = allCells.filter((cell) => typeof cell.formula === "string" && cell.formula.startsWith("="));
    const valueCells = allCells.filter((cell) => !cell.formula);

    const groupedFormulas = this.groupFormulaCells(formulaCells, filesMeta);
    const formulaProducerByCell = new Map<string, string>();
    for (const group of groupedFormulas) {
      for (const cell of group.cells) {
        formulaProducerByCell.set(toNodeId(group.fileName, group.sheet, cell), group.id);
      }
    }

    const referencedCells = new Set<string>();
    for (const group of groupedFormulas) {
      for (const ref of group.referenceDetails) {
        referencedCells.add(toNodeId(ref.file, ref.sheet, ref.cell));
      }
    }

    const groupedInputCandidates = valueCells.filter((cell) => {
      const cellId = toNodeId(cell.fileName, cell.sheet, cell.cell);
      return cell.fileRole === "input" || referencedCells.has(cellId);
    });

    const inputNodes: GraphNode[] = [];
    const inputProducerByCell = new Map<string, string>();
    const groupedInputs = this.groupInputCells(groupedInputCandidates);

    for (const group of groupedInputs) {
      const parsedRange = parseRangeRef(group.range);
      if (!parsedRange) {
        continue;
      }

      const nodeId = `input::${group.fileName}::${group.sheet}::${group.range}`;
      const valuesByCell = new Map(
        group.cells.map((entry) => [normalizeCellAddress(entry.cell), entry.value] as const)
      );
      const orderedValues = parsedRange.cells
        .map((cell) => valuesByCell.get(cell))
        .filter((value): value is NonNullable<typeof value> => value !== undefined);

      inputNodes.push({
        id: nodeId,
        nodeType: "input",
        fileName: group.fileName,
        fileRole: group.fileRole,
        sheet: group.sheet,
        range: group.range,
        shape: {
          rows: parsedRange.rows,
          cols: parsedRange.cols,
          size: parsedRange.size
        },
        operation: "ReadRange",
        inputs: [],
        output: {
          file: group.fileName,
          sheet: group.sheet,
          range: group.range
        },
        rangeValues: orderedValues,
        cell: parsedRange.startCell,
        value: orderedValues[0],
        dependencies: [],
        referenceDetails: []
      });

      for (const cell of parsedRange.cells) {
        inputProducerByCell.set(toNodeId(group.fileName, group.sheet, cell), nodeId);
      }
    }

    const nodeById = new Map<string, GraphNode>();
    for (const node of inputNodes) {
      nodeById.set(node.id, node);
    }

    const formulaNodes: GraphNode[] = groupedFormulas.map((group) => {
      const dependencies = new Set<string>();
      for (const ref of group.referenceDetails) {
        const refId = toNodeId(ref.file, ref.sheet, ref.cell);
        const producer = formulaProducerByCell.get(refId) ?? inputProducerByCell.get(refId);
        if (producer) {
          dependencies.add(producer);
        }
      }

      const parsedRange = parseRangeRef(group.outputRange);
      const shape = parsedRange
        ? { rows: parsedRange.rows, cols: parsedRange.cols, size: parsedRange.size }
        : { rows: 1, cols: 1, size: 1 };

      const dependencyArray = [...dependencies];
      const inputs = dependencyArray
        .map((depId) => nodeById.get(depId))
        .filter((dep): dep is GraphNode => Boolean(dep))
        .map((dep) => ({
          file: dep.fileName,
          sheet: dep.sheet,
          range: dep.range,
          nodeId: dep.id
        }));

      const node: GraphNode = {
        id: group.id,
        nodeType: "formula",
        fileName: group.fileName,
        fileRole: group.fileRole,
        sheet: group.sheet,
        range: group.outputRange,
        shape,
        operation: this.detectOperationName(group.formulaTemplate),
        inputs,
        output: {
          file: group.fileName,
          sheet: group.sheet,
          range: group.outputRange
        },
        formula: group.formulaTemplate,
        formulaByCell: group.formulaByCell,
        rangeValues: [],
        cell: group.anchorCell,
        dependencies: dependencyArray,
        referenceDetails: group.referenceDetails
      };

      nodeById.set(node.id, node);
      return node;
    });

    const outputNodes: GraphNode[] = formulaNodes.map((formulaNode) => {
      const outputId = `output::${formulaNode.fileName}::${formulaNode.sheet}::${formulaNode.range}`;
      const outputParsed = parseRangeRef(formulaNode.range);
      return {
        id: outputId,
        nodeType: "output",
        fileName: formulaNode.fileName,
        fileRole: formulaNode.fileRole,
        sheet: formulaNode.sheet,
        cell: outputParsed?.startCell ?? formulaNode.cell,
        range: formulaNode.range,
        shape: { ...formulaNode.shape },
        operation: "WriteRange",
        inputs: [
          {
            file: formulaNode.fileName,
            sheet: formulaNode.sheet,
            range: formulaNode.range,
            nodeId: formulaNode.id
          }
        ],
        output: {
          file: formulaNode.fileName,
          sheet: formulaNode.sheet,
          range: formulaNode.range
        },
        rangeValues: [],
        dependencies: [formulaNode.id],
        referenceDetails: []
      };
    });

    const nodes = [...inputNodes, ...formulaNodes, ...outputNodes];

    const uniqueSheets = new Set<string>();
    for (const workbook of workbooks) {
      for (const sheet of workbook.sheets) {
        uniqueSheets.add(`${workbook.fileName}::${sheet}`);
      }
    }

    return {
      nodes,
      edges: this.buildEdges(nodes),
      sheets: [...uniqueSheets],
      files: workbooks.map((workbook) => ({
        fileName: workbook.fileName,
        role: workbook.fileRole,
        sheets: workbook.sheets,
        uploadName: workbook.uploadName
      }))
    };
  }

  rebuildFromNodes(nodes: GraphNode[], files: BuildResult["files"]): BuildResult {
    const rebuiltNodes = nodes.map((node) => {
      const parsedRange = parseRangeRef(node.range);
      const shape = parsedRange
        ? { rows: parsedRange.rows, cols: parsedRange.cols, size: parsedRange.size }
        : node.shape;

      if (node.nodeType !== "formula") {
        return {
          ...node,
          shape
        };
      }

      const referenceDetails =
        node.referenceDetails.length > 0
          ? node.referenceDetails
          : this.formulaParserService.extractReferences(node.formula, node.sheet, node.fileName);

      return {
        ...node,
        shape,
        referenceDetails
      };
    });

    const uniqueSheets = new Set<string>();
    for (const node of rebuiltNodes) {
      uniqueSheets.add(`${node.fileName}::${node.sheet}`);
    }

    return {
      nodes: rebuiltNodes,
      edges: this.buildEdges(rebuiltNodes),
      sheets: [...uniqueSheets],
      files
    };
  }

  private buildEdges(nodes: GraphNode[]): GraphEdge[] {
    const edges: GraphEdge[] = [];
    for (const node of nodes) {
      for (const dep of node.dependencies) {
        edges.push({
          source: dep,
          target: node.id
        });
      }
    }
    return edges;
  }

  private groupFormulaCells(
    cells: ParsedCell[],
    filesMeta: Array<{ fileName: string; role: "input" | "output" | "other" }>
  ): Array<{
    id: string;
    fileName: string;
    fileRole: "input" | "output" | "other";
    sheet: string;
    anchorCell: string;
    outputRange: string;
    formulaTemplate: string;
    formulaByCell: Record<string, string>;
    referenceDetails: CellReference[];
    cells: string[];
  }> {
    const bySignature = new Map<string, ParsedCell[]>();

    for (const cell of cells) {
      const signature = this.computeFormulaSignature(cell);
      if (!signature) {
        continue;
      }
      if (!bySignature.has(signature)) {
        bySignature.set(signature, []);
      }
      bySignature.get(signature)?.push(cell);
    }

    const groups: Array<{
      id: string;
      fileName: string;
      fileRole: "input" | "output" | "other";
      sheet: string;
      anchorCell: string;
      outputRange: string;
      formulaTemplate: string;
      formulaByCell: Record<string, string>;
      referenceDetails: CellReference[];
      cells: string[];
    }> = [];

    for (const [_signature, candidates] of bySignature) {
      const components = this.connectedComponents(candidates);
      for (const component of components) {
        const range = this.tryRectangularRange(component.map((item) => item.cell));
        const chunks = range ? [component] : component.map((entry) => [entry]);

        for (const chunk of chunks) {
          const ordered = [...chunk].sort((left, right) => toNodeId(left.fileName, left.sheet, left.cell).localeCompare(toNodeId(right.fileName, right.sheet, right.cell)));
          const anchor = ordered[0];
          const rectangularRange = this.tryRectangularRange(ordered.map((item) => item.cell)) ?? encodeRange(anchor.cell);
          const refs = new Map<string, CellReference>();
          const formulaByCell: Record<string, string> = {};

          for (const item of ordered) {
            if (item.formula) {
              formulaByCell[normalizeCellAddress(item.cell)] = item.formula;
            }
            const extracted = this.formulaParserService.extractReferences(item.formula, item.sheet, item.fileName);
            for (const ref of extracted) {
              const resolved = ref.external
                ? { ...ref, file: this.resolveExternalFileToken(ref.file, item.fileName, filesMeta) }
                : ref;
              refs.set(`${resolved.file}::${resolved.sheet}::${resolved.cell}`, resolved);
            }
          }

          groups.push({
            id: `formula::${anchor.fileName}::${anchor.sheet}::${rectangularRange}`,
            fileName: anchor.fileName,
            fileRole: anchor.fileRole,
            sheet: anchor.sheet,
            anchorCell: normalizeCellAddress(anchor.cell),
            outputRange: rectangularRange,
            formulaTemplate: anchor.formula ?? "",
            formulaByCell,
            referenceDetails: [...refs.values()],
            cells: ordered.map((item) => normalizeCellAddress(item.cell))
          });
        }
      }
    }

    return groups;
  }

  private connectedComponents(cells: ParsedCell[]): ParsedCell[][] {
    const components: ParsedCell[][] = [];
    const byKey = new Map<string, ParsedCell>();
    for (const cell of cells) {
      byKey.set(toNodeId(cell.fileName, cell.sheet, cell.cell), cell);
    }

    const visited = new Set<string>();
    const deltas: Array<[number, number]> = [
      [1, 0],
      [-1, 0],
      [0, 1],
      [0, -1]
    ];

    for (const cell of cells) {
      const startId = toNodeId(cell.fileName, cell.sheet, cell.cell);
      if (visited.has(startId)) {
        continue;
      }

      const queue = [cell];
      const bucket: ParsedCell[] = [];
      visited.add(startId);

      while (queue.length > 0) {
        const current = queue.shift()!;
        bucket.push(current);
        const parsed = parseCellRef(current.cell);
        if (!parsed) {
          continue;
        }

        for (const [dc, dr] of deltas) {
          const colNum = colToNumber(parsed.col) + dc;
          const rowNum = parsed.row + dr;
          if (colNum < 1 || rowNum < 1) {
            continue;
          }

          const neighborCell = `${numberToCol(colNum)}${rowNum}`;
          const neighborId = toNodeId(current.fileName, current.sheet, neighborCell);
          const neighbor = byKey.get(neighborId);
          if (!neighbor || visited.has(neighborId)) {
            continue;
          }

          visited.add(neighborId);
          queue.push(neighbor);
        }
      }

      components.push(bucket);
    }

    return components;
  }

  private tryRectangularRange(cells: string[]): string | null {
    const parsed = cells.map((cell) => parseCellRef(cell)).filter((item): item is NonNullable<typeof item> => Boolean(item));
    if (parsed.length !== cells.length || parsed.length === 0) {
      return null;
    }

    const rows = parsed.map((item) => item.row);
    const cols = parsed.map((item) => {
      let n = 0;
      for (let i = 0; i < item.col.length; i += 1) {
        n = n * 26 + (item.col.charCodeAt(i) - 64);
      }
      return n;
    });

    const minRow = Math.min(...rows);
    const maxRow = Math.max(...rows);
    const minCol = Math.min(...cols);
    const maxCol = Math.max(...cols);
    const area = (maxRow - minRow + 1) * (maxCol - minCol + 1);
    if (area !== cells.length) {
      return null;
    }

    const toCol = (num: number): string => {
      let n = num;
      let col = "";
      while (n > 0) {
        const rem = (n - 1) % 26;
        col = String.fromCharCode(65 + rem) + col;
        n = Math.floor((n - 1) / 26);
      }
      return col;
    };

    return encodeRange(`${toCol(minCol)}${minRow}`, `${toCol(maxCol)}${maxRow}`);
  }

  private groupInputCells(cells: ParsedCell[]): Array<{
    fileName: string;
    fileRole: "input" | "output" | "other";
    sheet: string;
    range: string;
    cells: Array<{ cell: string; value?: string | number | boolean }>;
  }> {
    const bySheet = new Map<string, ParsedCell[]>();
    for (const cell of cells) {
      const key = `${cell.fileName}::${cell.sheet}`;
      if (!bySheet.has(key)) {
        bySheet.set(key, []);
      }
      bySheet.get(key)?.push(cell);
    }

    const result: Array<{
      fileName: string;
      fileRole: "input" | "output" | "other";
      sheet: string;
      range: string;
      cells: Array<{ cell: string; value?: string | number | boolean }>;
    }> = [];

    for (const [, sheetCells] of bySheet) {
      const components = this.connectedComponents(sheetCells);
      for (const component of components) {
        const normalizedCells = component
          .map((entry) => ({
            ...entry,
            cell: normalizeCellAddress(entry.cell)
          }))
          .sort((left, right) => toNodeId(left.fileName, left.sheet, left.cell).localeCompare(toNodeId(right.fileName, right.sheet, right.cell)));

        if (normalizedCells.length === 0) {
          continue;
        }

        const rectangularRange = this.tryRectangularRange(normalizedCells.map((entry) => entry.cell));
        if (rectangularRange) {
          result.push({
            fileName: normalizedCells[0].fileName,
            fileRole: normalizedCells[0].fileRole,
            sheet: normalizedCells[0].sheet,
            range: rectangularRange,
            cells: normalizedCells.map((entry) => ({
              cell: entry.cell,
              value: entry.value
            }))
          });
          continue;
        }

        const byColumn = new Map<string, Array<{ cell: string; value?: string | number | boolean; row: number }>>();
        for (const entry of normalizedCells) {
          const parsed = parseCellRef(entry.cell);
          if (!parsed) {
            continue;
          }
          if (!byColumn.has(parsed.col)) {
            byColumn.set(parsed.col, []);
          }
          byColumn.get(parsed.col)?.push({
            cell: entry.cell,
            value: entry.value,
            row: parsed.row
          });
        }

        for (const [col, colCells] of byColumn) {
          const sorted = [...colCells].sort((left, right) => left.row - right.row);
          let runStart = 0;

          for (let index = 1; index <= sorted.length; index += 1) {
            const isBreak = index === sorted.length || sorted[index].row !== sorted[index - 1].row + 1;
            if (!isBreak) {
              continue;
            }

            const run = sorted.slice(runStart, index);
            const start = run[0].row;
            const end = run[run.length - 1].row;
            result.push({
              fileName: normalizedCells[0].fileName,
              fileRole: normalizedCells[0].fileRole,
              sheet: normalizedCells[0].sheet,
              range: encodeRange(`${col}${start}`, `${col}${end}`),
              cells: run.map((entry) => ({
                cell: entry.cell,
                value: entry.value
              }))
            });

            runStart = index;
          }
        }
      }
    }

    return result;
  }

  private detectOperationName(formula: string): string {
    const body = formula.trim().replace(/^=/, "").toUpperCase();
    if (body.includes("^2") || /POWER\([^,]+,\s*2\)/.test(body)) {
      return "Square";
    }
    if (body.startsWith("SUM(")) {
      return "Sum";
    }
    return "ExcelFormula";
  }

  private computeFormulaSignature(cell: ParsedCell): string | null {
    const formula = cell.formula?.trim();
    if (!formula || !formula.startsWith("=")) {
      return null;
    }

    const anchor = parseCellRef(cell.cell);
    if (!anchor) {
      return null;
    }

    const normalized = formula.slice(1).replace(GraphBuilderService.FORMULA_TOKEN_REGEX, (matched, t1: string, _r: string, t2?: string) => {
      const p1 = this.parseTokenPattern(t1, cell.fileName, cell.sheet, anchor.col, anchor.row);
      if (!p1) {
        return matched;
      }
      if (!t2) {
        return `R(${p1})`;
      }
      const p2 = this.parseTokenPattern(t2, cell.fileName, cell.sheet, anchor.col, anchor.row);
      return p2 ? `RG(${p1}|${p2})` : `R(${p1})`;
    });

    return `${cell.fileName}::${cell.sheet}::${normalized.replace(/\s+/g, "").toUpperCase()}`;
  }

  private parseTokenPattern(
    token: string,
    fileName: string,
    sheet: string,
    anchorCol: string,
    anchorRow: number
  ): string | null {
    const clean = token.trim().replace(/\$/g, "");
    const [sheetPart, cellPartMaybe] = clean.includes("!") ? clean.split("!") : [null, clean];
    const cellPart = cellPartMaybe ?? clean;
    const parsed = parseCellRef(cellPart);
    if (!parsed) {
      return null;
    }

    const colToNum = (col: string): number => {
      let n = 0;
      for (let i = 0; i < col.length; i += 1) {
        n = n * 26 + (col.charCodeAt(i) - 64);
      }
      return n;
    };

    const refToken = sheetPart?.replace(/^'|'$/g, "") ?? sheet;
    const externalMatch = refToken.match(/^(?:.*[\\/])?\[([^\]]+)\](.+)$/);
    const refFile = externalMatch ? externalMatch[1].trim() : fileName;
    const refSheet = externalMatch ? externalMatch[2].trim() : refToken;
    const scope = refFile === fileName ? "local" : `ext:${refFile.toUpperCase()}`;

    return `${scope}:${refSheet.toUpperCase()}:${colToNum(parsed.col) - colToNum(anchorCol)}:${parsed.row - anchorRow}`;
  }
}

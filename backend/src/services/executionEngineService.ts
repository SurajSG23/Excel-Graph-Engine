import * as formulajs from "formulajs";
import { CellValue, GraphNode, ValidationIssue } from "../models/graph";
import { normalizeCellAddress, normalizeFileName, normalizeSheetName, parseRangeRef } from "../utils/cellUtils";

interface RecomputeResult {
  nodes: GraphNode[];
  issues: ValidationIssue[];
}

export class ExecutionEngineService {
  recompute(nodes: GraphNode[], changedNodeIds: string[] = []): RecomputeResult {
    const nodeMap = new Map(nodes.map((node) => [node.id, { ...node }]));
    const nodeIds = new Set(nodeMap.keys());

    const dependentsMap = new Map<string, string[]>();
    for (const node of nodeMap.values()) {
      for (const dep of node.dependencies) {
        if (!dependentsMap.has(dep)) {
          dependentsMap.set(dep, []);
        }
        dependentsMap.get(dep)?.push(node.id);
      }
    }

    const affected = this.computeAffectedNodes(changedNodeIds, dependentsMap);
    const evalSet = affected.size > 0 ? affected : new Set(nodeMap.keys());
    const order = this.topologicalSort([...nodeMap.values()], nodeIds);

    const issues: ValidationIssue[] = [];
    const cellValues = this.seedCellValues([...nodeMap.values()]);

    for (const nodeId of order) {
      if (!evalSet.has(nodeId)) {
        continue;
      }

      const node = nodeMap.get(nodeId);
      if (!node) {
        continue;
      }

      if (node.nodeType === "input") {
        const values = node.rangeValues ?? (node.value !== undefined ? [node.value] : []);
        node.rangeValues = values;
        if (values.length > 0) {
          this.writeRangeValues(cellValues, node.fileName, node.sheet, node.range, values);
          node.value = values[0];
        }
        nodeMap.set(node.id, node);
        continue;
      }

      if (node.nodeType === "formula") {
        const result = this.evaluateFormulaNode(node, cellValues);
        if (result.error) {
          issues.push({
            type: "INVALID_FORMULA",
            nodeId: node.id,
            message: `Failed to evaluate ${node.id}: ${result.error}`
          });
          continue;
        }

        node.rangeValues = result.values;
        node.value = result.values[0];
        this.writeRangeValues(cellValues, node.fileName, node.sheet, node.range, result.values);
        nodeMap.set(node.id, node);
        continue;
      }

      const sourceNode = node.dependencies.length > 0 ? nodeMap.get(node.dependencies[0]) : undefined;
      const outputValues = sourceNode?.rangeValues ?? [];
      node.rangeValues = [...outputValues];
      node.value = outputValues[0];
      this.writeRangeValues(cellValues, node.fileName, node.sheet, node.range, node.rangeValues);
      nodeMap.set(node.id, node);
    }

    return {
      nodes: [...nodeMap.values()],
      issues
    };
  }

  private seedCellValues(nodes: GraphNode[]): Map<string, CellValue> {
    const values = new Map<string, CellValue>();

    for (const node of nodes) {
      if (node.nodeType === "input") {
        this.writeRangeValues(values, node.fileName, node.sheet, node.range, node.rangeValues ?? []);
      }
    }

    return values;
  }

  private evaluateFormulaNode(
    node: GraphNode,
    cellValues: Map<string, CellValue>
  ): { values: CellValue[]; error?: string } {
    try {
      if (node.formulaByCell && Object.keys(node.formulaByCell).length > 0) {
        const range = parseRangeRef(node.range);
        if (!range) {
          return { values: [], error: `Invalid range ${node.range}` };
        }

        const byCell = new Map<string, CellValue>();
        for (const [cell, formula] of Object.entries(node.formulaByCell)) {
          const evaluated = this.evaluateFormula(formula, node.fileName, node.sheet, cellValues);
          if (evaluated.error) {
            return { values: [], error: evaluated.error };
          }
          if (evaluated.value !== undefined) {
            byCell.set(normalizeCellAddress(cell), evaluated.value);
          }
        }

        return {
          values: range.cells.map((cell) => byCell.get(cell) ?? "")
        };
      }

      if (node.formula) {
        const evaluated = this.evaluateFormula(node.formula, node.fileName, node.sheet, cellValues);
        if (evaluated.error) {
          return { values: [], error: evaluated.error };
        }
        return {
          values: evaluated.value !== undefined ? [evaluated.value] : []
        };
      }

      if (node.operation === "Square") {
        const flat = node.inputs.flatMap((input) => this.readRangeValues(cellValues, input.file, input.sheet, input.range));
        return {
          values: flat.map((value) => (typeof value === "number" ? value * value : value))
        };
      }

      const passthrough = node.inputs.flatMap((input) => this.readRangeValues(cellValues, input.file, input.sheet, input.range));
      return { values: passthrough };
    } catch (error) {
      return {
        values: [],
        error: error instanceof Error ? error.message : "Unknown formula evaluation error"
      };
    }
  }

  private readRangeValues(
    cellValues: Map<string, CellValue>,
    fileName: string,
    sheet: string,
    rangeRef: string
  ): CellValue[] {
    const parsed = parseRangeRef(rangeRef);
    if (!parsed) {
      return [];
    }

    return parsed.cells.map((cell) => cellValues.get(this.cellKey(fileName, sheet, cell)) ?? "");
  }

  private writeRangeValues(
    cellValues: Map<string, CellValue>,
    fileName: string,
    sheet: string,
    rangeRef: string,
    values: CellValue[]
  ): void {
    const parsed = parseRangeRef(rangeRef);
    if (!parsed || values.length === 0) {
      return;
    }

    for (let i = 0; i < parsed.cells.length && i < values.length; i += 1) {
      cellValues.set(this.cellKey(fileName, sheet, parsed.cells[i]), values[i]);
    }
  }

  private cellKey(fileName: string, sheet: string, cell: string): string {
    return `${normalizeFileName(fileName)}::${normalizeSheetName(sheet)}::${normalizeCellAddress(cell)}`;
  }

  private computeAffectedNodes(changedNodeIds: string[], dependentsMap: Map<string, string[]>): Set<string> {
    const affected = new Set<string>();
    const queue = [...changedNodeIds];

    while (queue.length > 0) {
      const current = queue.shift();
      if (!current || affected.has(current)) {
        continue;
      }

      affected.add(current);
      for (const dependent of dependentsMap.get(current) ?? []) {
        queue.push(dependent);
      }
    }

    return affected;
  }

  private topologicalSort(nodes: GraphNode[], nodeIds: Set<string>): string[] {
    const indegree = new Map<string, number>();
    const outgoing = new Map<string, string[]>();

    for (const node of nodes) {
      indegree.set(node.id, 0);
      outgoing.set(node.id, []);
    }

    for (const node of nodes) {
      for (const dep of node.dependencies) {
        if (!nodeIds.has(dep)) {
          continue;
        }
        indegree.set(node.id, (indegree.get(node.id) ?? 0) + 1);
        outgoing.get(dep)?.push(node.id);
      }
    }

    const queue = [...nodes.filter((node) => (indegree.get(node.id) ?? 0) === 0).map((node) => node.id)];
    const order: string[] = [];

    while (queue.length > 0) {
      const current = queue.shift()!;
      order.push(current);

      for (const next of outgoing.get(current) ?? []) {
        const nextDegree = (indegree.get(next) ?? 0) - 1;
        indegree.set(next, nextDegree);
        if (nextDegree === 0) {
          queue.push(next);
        }
      }
    }

    if (order.length < nodes.length) {
      for (const node of nodes) {
        if (!order.includes(node.id)) {
          order.push(node.id);
        }
      }
    }

    return order;
  }

  private evaluateFormula(
    formula: string,
    currentFileName: string,
    currentSheet: string,
    cellValues: Map<string, CellValue>
  ): { value: CellValue | undefined; error?: string } {
    try {
      const expression = this.transformFormulaToJs(formula, currentFileName, currentSheet, cellValues);
      const helpers = {
        fn: (name: string) => {
          const upper = name.toUpperCase();
          return (...args: unknown[]) => {
            const flat = args.flatMap((arg) => (Array.isArray(arg) ? arg : [arg]));
            const fn = (formulajs as Record<string, (...innerArgs: unknown[]) => unknown>)[upper];
            if (!fn) {
              throw new Error(`Unsupported function: ${name}`);
            }
            return fn(...flat);
          };
        }
      };

      // eslint-disable-next-line no-new-func
      const fn = new Function("helpers", `return (${expression});`) as (h: typeof helpers) => unknown;
      const raw = fn(helpers);

      if (typeof raw === "number" || typeof raw === "string" || typeof raw === "boolean") {
        return { value: raw };
      }

      return { value: undefined };
    } catch (error) {
      return {
        value: undefined,
        error: error instanceof Error ? error.message : "Unknown formula evaluation error"
      };
    }
  }

  private transformFormulaToJs(
    formula: string,
    currentFileName: string,
    currentSheet: string,
    cellValues: Map<string, CellValue>
  ): string {
    let body = formula.startsWith("=") ? formula.slice(1) : formula;

    const rangeRegex = /((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+):((?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g;
    body = body.replace(rangeRegex, (_m, left, right) => {
      const leftParsed = this.parseRef(left, currentFileName, currentSheet);
      const rightParsed = this.parseRef(right, currentFileName, currentSheet);

      if (leftParsed.file !== rightParsed.file || leftParsed.sheet !== rightParsed.sheet) {
        const leftValue = this.getNumericValue(cellValues.get(this.cellKey(leftParsed.file, leftParsed.sheet, leftParsed.cell)));
        const rightValue = this.getNumericValue(cellValues.get(this.cellKey(rightParsed.file, rightParsed.sheet, rightParsed.cell)));
        return `[${leftValue},${rightValue}]`;
      }

      const cells = parseRangeRef(`${leftParsed.cell}:${rightParsed.cell}`)?.cells ?? [];
      const values = cells.map((cell) => this.getNumericValue(cellValues.get(this.cellKey(leftParsed.file, leftParsed.sheet, cell))));
      return `[${values.join(",")}]`;
    });

    const referenceRegex = /(?:'[^']+'|\[[^\]]+\][^!]+|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+/g;
    body = body.replace(referenceRegex, (token) => {
      const parsed = this.parseRef(token, currentFileName, currentSheet);
      const value = cellValues.get(this.cellKey(parsed.file, parsed.sheet, parsed.cell));
      return String(this.getNumericValue(value));
    });

    body = body.replace(/#(REF!|N\/A|DIV\/0!|VALUE!|NAME\?|NUM!|NULL!)/gi, "0");
    body = body.replace(/\^/g, "**");
    body = body.replace(/\b([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g, (_match, name: string) => {
      const upper = name.toUpperCase();
      if (["TRUE", "FALSE"].includes(upper)) {
        return `${upper === "TRUE" ? "true" : "false"}(`;
      }
      return `helpers.fn(\"${upper}\")(`;
    });

    return body;
  }

  private parseRef(
    ref: string,
    currentFileName: string,
    currentSheet: string
  ): { file: string; sheet: string; cell: string } {
    const cleaned = ref.trim();
    if (cleaned.includes("!")) {
      const [rawSheet, rawCell] = cleaned.split("!");
      const sheetToken = rawSheet.replace(/^'|'$/g, "");
      const external = sheetToken.match(/^(?:.*[\\/])?\[([^\]]+)\](.+)$/);
      return {
        file: normalizeFileName(external ? external[1] : currentFileName),
        sheet: normalizeSheetName(external ? external[2] : sheetToken),
        cell: normalizeCellAddress(rawCell)
      };
    }

    return {
      file: normalizeFileName(currentFileName),
      sheet: currentSheet,
      cell: normalizeCellAddress(cleaned)
    };
  }

  private getNumericValue(value: CellValue | undefined): number {
    return typeof value === "number" && Number.isFinite(value) ? value : 0;
  }
}

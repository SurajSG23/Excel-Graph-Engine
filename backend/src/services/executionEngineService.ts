import * as formulajs from "formulajs";
import { GraphNode, ValidationIssue } from "../models/graph";

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
        dependentsMap.get(dep)!.push(node.id);
      }
    }

    const affected = this.computeAffectedNodes(changedNodeIds, dependentsMap);
    const evalSet = affected.size > 0 ? affected : new Set(nodeMap.keys());

    const order = this.topologicalSort([...nodeMap.values()], nodeIds);
    const issues: ValidationIssue[] = [];

    for (const nodeId of order) {
      if (!evalSet.has(nodeId)) {
        continue;
      }

      const node = nodeMap.get(nodeId);
      if (!node || !node.formula) {
        continue;
      }

      const result = this.evaluateFormula(node.formula, node.sheet, nodeMap);
      if (result.error) {
        issues.push({
          type: "INVALID_FORMULA",
          nodeId,
          message: `Failed to evaluate ${nodeId}: ${result.error}`
        });
        continue;
      }

      node.value = result.value;
      nodeMap.set(nodeId, node);
    }

    return {
      nodes: [...nodeMap.values()],
      issues
    };
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

    // Cycles are validated elsewhere; append leftovers to keep deterministic behavior.
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
    currentSheet: string,
    nodeMap: Map<string, GraphNode>
  ): { value: number | undefined; error?: string } {
    try {
      const expression = this.transformFormulaToJs(formula, currentSheet, nodeMap);
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

      if (typeof raw === "number") {
        return { value: raw };
      }

      if (typeof raw === "string" && raw.trim() !== "") {
        const num = Number(raw);
        return Number.isFinite(num) ? { value: num } : { value: undefined };
      }

      return { value: undefined };
    } catch (error) {
      return {
        value: undefined,
        error: error instanceof Error ? error.message : "Unknown formula evaluation error"
      };
    }
  }

  private transformFormulaToJs(formula: string, currentSheet: string, nodeMap: Map<string, GraphNode>): string {
    let body = formula.startsWith("=") ? formula.slice(1) : formula;

    const rangeRegex = /((?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+):((?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g;
    body = body.replace(rangeRegex, (_m, left, right) => {
      const leftParsed = this.parseRef(left, currentSheet);
      const rightParsed = this.parseRef(right, currentSheet);
      const cells = this.expandRangeForEvaluation(leftParsed.sheet, leftParsed.cell, rightParsed.sheet, rightParsed.cell);
      const values = cells.map((id) => this.getNodeNumericValue(id, nodeMap));
      return `[${values.join(",")}]`;
    });

    const referenceRegex = /(?:'[^']+'|[A-Za-z0-9_]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+/g;
    body = body.replace(referenceRegex, (token) => {
      const parsed = this.parseRef(token, currentSheet);
      const id = `${parsed.sheet}!${parsed.cell}`;
      const value = this.getNodeNumericValue(id, nodeMap);
      return String(value);
    });

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

  private parseRef(ref: string, currentSheet: string): { sheet: string; cell: string } {
    const cleaned = ref.replace(/\$/g, "").trim();
    if (cleaned.includes("!")) {
      const [sheet, cell] = cleaned.split("!");
      return {
        sheet: sheet.replace(/^'|'$/g, ""),
        cell: cell.toUpperCase()
      };
    }

    return {
      sheet: currentSheet,
      cell: cleaned.toUpperCase()
    };
  }

  private expandRangeForEvaluation(
    startSheet: string,
    startCell: string,
    endSheet: string,
    endCell: string
  ): string[] {
    if (startSheet !== endSheet) {
      return [`${startSheet}!${startCell}`, `${endSheet}!${endCell}`];
    }

    const [startCol, startRow] = this.splitCell(startCell);
    const [endCol, endRow] = this.splitCell(endCell);
    if (!startCol || !endCol || !startRow || !endRow) {
      return [`${startSheet}!${startCell}`, `${endSheet}!${endCell}`];
    }

    const startColNum = this.colToNum(startCol);
    const endColNum = this.colToNum(endCol);
    const minCol = Math.min(startColNum, endColNum);
    const maxCol = Math.max(startColNum, endColNum);
    const minRow = Math.min(Number(startRow), Number(endRow));
    const maxRow = Math.max(Number(startRow), Number(endRow));

    const refs: string[] = [];
    for (let col = minCol; col <= maxCol; col += 1) {
      for (let row = minRow; row <= maxRow; row += 1) {
        refs.push(`${startSheet}!${this.numToCol(col)}${row}`);
      }
    }

    return refs;
  }

  private getNodeNumericValue(id: string, nodeMap: Map<string, GraphNode>): number {
    const value = nodeMap.get(id)?.value;
    return typeof value === "number" && Number.isFinite(value) ? value : 0;
  }

  private splitCell(cell: string): [string | null, string | null] {
    const match = cell.match(/^([A-Z]{1,3})([0-9]+)$/);
    if (!match) {
      return [null, null];
    }
    return [match[1], match[2]];
  }

  private colToNum(col: string): number {
    let n = 0;
    for (let i = 0; i < col.length; i += 1) {
      n = n * 26 + (col.charCodeAt(i) - 64);
    }
    return n;
  }

  private numToCol(num: number): string {
    let n = num;
    let col = "";
    while (n > 0) {
      const rem = (n - 1) % 26;
      col = String.fromCharCode(65 + rem) + col;
      n = Math.floor((n - 1) / 26);
    }
    return col;
  }
}

import {
  FormulaNodeConfig,
  ParsedWorkbookData,
  PipelineConfig,
  PipelineGraph,
  PipelineGraphEdge,
  PipelineGraphNode
} from "../models/pipeline";
import { FormulaGrouper } from "./formula_grouper";
import { collapseCellsToRange, expandRange, rangesToCellKeys } from "./node_models";

export interface BuildPipelineResult {
  config: PipelineConfig;
  graph: PipelineGraph;
  executionOrder: string[];
  dependencies: Map<string, string[]>;
}

function toNodeOutputs(node: FormulaNodeConfig): Set<string> {
  const keys = new Set<string>();
  for (const cell of node.outputCells.length > 0 ? node.outputCells : expandRange(node.output.range)) {
    keys.add(`${node.output.sheet}!${cell}`);
  }
  return keys;
}

function buildDependencies(formulas: FormulaNodeConfig[]): Map<string, string[]> {
  const outputsByNode = new Map(formulas.map((node) => [node.id, toNodeOutputs(node)]));
  const dependencies = new Map<string, string[]>();

  for (const node of formulas) {
    const inputCells = rangesToCellKeys(node.inputs);
    const deps: string[] = [];

    for (const candidate of formulas) {
      if (candidate.id === node.id) {
        continue;
      }

      const candidateOutputs = outputsByNode.get(candidate.id) ?? new Set<string>();
      const touches = [...candidateOutputs].some((cellKey) => inputCells.has(cellKey));
      if (touches) {
        deps.push(candidate.id);
      }
    }

    dependencies.set(node.id, deps);
  }

  return dependencies;
}

function topologicalOrder(formulas: FormulaNodeConfig[], deps: Map<string, string[]>): string[] {
  const indegree = new Map<string, number>();
  const outgoing = new Map<string, string[]>();

  for (const node of formulas) {
    indegree.set(node.id, 0);
    outgoing.set(node.id, []);
  }

  for (const node of formulas) {
    for (const dep of deps.get(node.id) ?? []) {
      indegree.set(node.id, (indegree.get(node.id) ?? 0) + 1);
      outgoing.get(dep)?.push(node.id);
    }
  }

  const queue = formulas
    .map((node) => node.id)
    .filter((id) => (indegree.get(id) ?? 0) === 0)
    .sort((a, b) => a.localeCompare(b));

  const order: string[] = [];
  while (queue.length > 0) {
    const current = queue.shift()!;
    order.push(current);
    for (const next of outgoing.get(current) ?? []) {
      const deg = (indegree.get(next) ?? 0) - 1;
      indegree.set(next, deg);
      if (deg === 0) {
        queue.push(next);
      }
    }
    queue.sort((a, b) => a.localeCompare(b));
  }

  if (order.length < formulas.length) {
    const seen = new Set(order);
    const remaining = formulas
      .map((node) => node.id)
      .filter((id) => !seen.has(id))
      .sort((a, b) => a.localeCompare(b));
    order.push(...remaining);
  }

  return order;
}

export class PipelineBuilder {
  private readonly formulaGrouper = new FormulaGrouper();

  build(parsed: ParsedWorkbookData): BuildPipelineResult {
    const formulas = this.formulaGrouper.group(parsed.formulaCells);
    const allInputRefs = formulas.flatMap((node) =>
      node.inputs.flatMap((item) => expandRange(item.range).map((cell) => ({ sheet: item.sheet, cell })))
    );

    const inputRanges = (() => {
      const bySheet = new Map<string, Set<string>>();
      for (const ref of allInputRefs) {
        if (!bySheet.has(ref.sheet)) {
          bySheet.set(ref.sheet, new Set());
        }
        bySheet.get(ref.sheet)!.add(ref.cell);
      }
      return [...bySheet.entries()].map(([sheet, cells]) => ({
        sheet,
        range: collapseCellsToRange([...cells])
      }));
    })();

    const outputRanges = formulas.map((node) => ({ ...node.output }));

    const config: PipelineConfig = {
      input: {
        id: "input",
        name: "Input",
        filePath: parsed.sourceFilePath,
        sheets: parsed.sheetNames,
        ranges: inputRanges
      },
      formulas,
      output: {
        id: "output",
        name: "Output",
        targetFilePath: parsed.targetFilePath,
        ranges: outputRanges
      }
    };

    const nodes: PipelineGraphNode[] = [
      { id: "input", type: "input", label: "Input" },
      ...formulas.map((node) => ({ id: node.id, type: "formula" as const, label: node.name })),
      { id: "output", type: "output", label: "Output" }
    ];

    const edges: PipelineGraphEdge[] = [
      ...formulas.map((node) => ({ source: "input", target: node.id })),
      ...formulas.map((node) => ({ source: node.id, target: "output" }))
    ];

    const dependencies = buildDependencies(formulas);
    const executionOrder = topologicalOrder(formulas, dependencies);

    return {
      config,
      graph: { nodes, edges },
      executionOrder,
      dependencies
    };
  }

  rebuild(config: PipelineConfig): BuildPipelineResult {
    const formulas = config.formulas;
    const dependencies = buildDependencies(formulas);
    const executionOrder = topologicalOrder(formulas, dependencies);
    const graph: PipelineGraph = {
      nodes: [
        { id: "input", type: "input", label: "Input" },
        ...formulas.map((node) => ({ id: node.id, type: "formula" as const, label: node.name })),
        { id: "output", type: "output", label: "Output" }
      ],
      edges: [
        ...formulas.map((node) => ({ source: "input", target: node.id })),
        ...formulas.map((node) => ({ source: node.id, target: "output" }))
      ]
    };

    return {
      config,
      graph,
      executionOrder,
      dependencies
    };
  }
}

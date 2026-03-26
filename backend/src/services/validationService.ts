import { GraphNode, ValidationIssue } from "../models/graph";
import { parseRangeRef, rangesOverlap } from "../utils/cellUtils";

/**
 * ValidationService validates the range-based graph for consistency and correctness.
 * All validation operates at the range level, not cell level.
 */
export class ValidationService {
  validate(
    nodes: GraphNode[],
    files: Array<{ fileName: string; sheets: string[] }> = []
  ): ValidationIssue[] {
    const issues: ValidationIssue[] = [];
    const nodeSet = new Set(nodes.map((node) => node.id));
    const fileSet = new Set(files.map((file) => file.fileName));
    const sheetSet = new Set(files.flatMap((file) => file.sheets.map((sheet) => `${file.fileName}::${sheet}`)));

    for (const node of nodes) {
      if (!parseRangeRef(node.range)) {
        issues.push({
          type: "INVALID_RANGE",
          nodeId: node.id,
          message: `Invalid range '${node.range}' on node ${node.id}`,
          relatedNodeIds: [node.id]
        });
      }

      if (node.nodeType === "formula") {
        if (node.formula && !node.formula.startsWith("=")) {
          issues.push({
            type: "INVALID_FORMULA",
            nodeId: node.id,
            message: `Invalid formula in ${node.id}. Formulas must start with '='.`
          });
        }

        if (node.shape.size > 0 && node.rangeValues && node.rangeValues.length > 0 && node.rangeValues.length !== node.shape.size) {
          issues.push({
            type: "MISMATCHED_RANGE_SHAPE",
            nodeId: node.id,
            message: `Node ${node.id} produced ${node.rangeValues.length} values for shape ${node.shape.rows}x${node.shape.cols}.`,
            relatedNodeIds: [node.id]
          });
        }
      }

      for (const input of node.inputs) {
        if (!parseRangeRef(input.range)) {
          issues.push({
            type: "INVALID_RANGE",
            nodeId: node.id,
            message: `Node ${node.id} has invalid input range '${input.range}'.`,
            relatedNodeIds: [node.id]
          });
        }

        if (fileSet.size > 0 && !fileSet.has(input.file)) {
          issues.push({
            type: "MISSING_REFERENCE",
            nodeId: node.id,
            message: `Node ${node.id} references missing workbook ${input.file}`,
            relatedNodeIds: [node.id]
          });
        }

        if (sheetSet.size > 0 && !sheetSet.has(`${input.file}::${input.sheet}`)) {
          issues.push({
            type: "MISSING_REFERENCE",
            nodeId: node.id,
            message: `Node ${node.id} references missing sheet ${input.sheet} in ${input.file}`,
            relatedNodeIds: [node.id]
          });
        }
      }

      for (const dep of node.dependencies) {
        if (!nodeSet.has(dep)) {
          issues.push({
            type: "MISSING_REFERENCE",
            nodeId: node.id,
            message: `Node ${node.id} depends on missing node ${dep}`,
            relatedNodeIds: [node.id, dep]
          });
        }
      }

      for (const ref of node.referenceDetails) {
        if (!ref.external) {
          continue;
        }

        if (fileSet.size > 0 && !fileSet.has(ref.file)) {
          issues.push({
            type: "MISSING_REFERENCE",
            nodeId: node.id,
            message: `Node ${node.id} references missing external workbook ${ref.file}`,
            relatedNodeIds: [node.id]
          });
          continue;
        }

        if (sheetSet.size > 0 && !sheetSet.has(`${ref.file}::${ref.sheet}`)) {
          issues.push({
            type: "MISSING_REFERENCE",
            nodeId: node.id,
            message: `Node ${node.id} references missing sheet ${ref.sheet} in ${ref.file}`,
            relatedNodeIds: [node.id]
          });
        }
      }
    }

    for (const overlap of this.detectOverlappingWrites(nodes)) {
      issues.push(overlap);
    }

    const circularPaths = this.detectCycles(nodes);
    for (const cycle of circularPaths) {
      issues.push({
        type: "CIRCULAR_DEPENDENCY",
        message: `Circular dependency detected: ${cycle.join(" -> ")}`,
        relatedNodeIds: cycle
      });
    }

    return issues;
  }

  /**
   * Detects overlapping write ranges between formula and output nodes.
   * Uses range-level overlap detection instead of cell-level expansion.
   */
  private detectOverlappingWrites(nodes: GraphNode[]): ValidationIssue[] {
    const issues: ValidationIssue[] = [];
    const writers = nodes.filter((node) => node.nodeType === "formula" || node.nodeType === "output");

    for (let i = 0; i < writers.length; i += 1) {
      const left = writers[i];
      if (!parseRangeRef(left.range)) {
        continue;
      }

      for (let j = i + 1; j < writers.length; j += 1) {
        const right = writers[j];
        if (left.fileName !== right.fileName || left.sheet !== right.sheet) {
          continue;
        }

        if (!parseRangeRef(right.range)) {
          continue;
        }

        // Use range-level overlap detection instead of cell expansion
        if (rangesOverlap(left.range, right.range)) {
          issues.push({
            type: "OVERLAPPING_WRITES",
            nodeId: left.id,
            message: `Overlapping write ranges detected between ${left.id} and ${right.id}.`,
            relatedNodeIds: [left.id, right.id]
          });
        }
      }
    }

    return issues;
  }

  private detectCycles(nodes: GraphNode[]): string[][] {
    const adjacency = new Map<string, string[]>();
    const nodeIds = new Set(nodes.map((node) => node.id));

    for (const node of nodes) {
      adjacency.set(
        node.id,
        node.dependencies.filter((dep) => nodeIds.has(dep))
      );
    }

    const result: string[][] = [];
    const state = new Map<string, "white" | "gray" | "black">();
    const stack: string[] = [];

    const dfs = (nodeId: string): void => {
      state.set(nodeId, "gray");
      stack.push(nodeId);

      for (const dep of adjacency.get(nodeId) ?? []) {
        const depState = state.get(dep) ?? "white";
        if (depState === "white") {
          dfs(dep);
        } else if (depState === "gray") {
          const startIdx = stack.indexOf(dep);
          if (startIdx >= 0) {
            result.push([...stack.slice(startIdx), dep]);
          }
        }
      }

      stack.pop();
      state.set(nodeId, "black");
    };

    for (const id of adjacency.keys()) {
      if ((state.get(id) ?? "white") === "white") {
        dfs(id);
      }
    }

    return result;
  }
}

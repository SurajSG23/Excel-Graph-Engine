import { GraphNode, ValidationIssue } from "../models/graph";

export class ValidationService {
  validate(nodes: GraphNode[]): ValidationIssue[] {
    const issues: ValidationIssue[] = [];
    const nodeSet = new Set(nodes.map((node) => node.id));

    for (const node of nodes) {
      for (const dep of node.dependencies) {
        if (!nodeSet.has(dep)) {
          issues.push({
            type: "MISSING_REFERENCE",
            nodeId: node.id,
            message: `Node ${node.id} references missing cell ${dep}`,
            relatedNodeIds: [node.id, dep]
          });
        }
      }

      if (node.formula && !node.formula.startsWith("=")) {
        issues.push({
          type: "INVALID_FORMULA",
          nodeId: node.id,
          message: `Invalid formula in ${node.id}. Formulas must start with '='.`
        });
      }
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

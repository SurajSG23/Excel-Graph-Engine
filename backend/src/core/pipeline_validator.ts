import { PipelineConfig, ValidationIssue } from "../models/pipeline";
import { expandRange, rangesAreValid } from "./node_models";

export class PipelineValidator {
  validate(config: PipelineConfig, executionOrder: string[]): ValidationIssue[] {
    const issues: ValidationIssue[] = [];

    for (const node of config.formulas) {
      if (!node.formula.trim().startsWith("=")) {
        issues.push({
          type: "INVALID_FORMULA",
          nodeId: node.id,
          message: `${node.name} formula must start with '='.`
        });
      }

      if (!rangesAreValid(node.inputs)) {
        issues.push({
          type: "INVALID_RANGE",
          nodeId: node.id,
          message: `${node.name} has one or more invalid input ranges.`
        });
      }

      if (!rangesAreValid([node.output])) {
        issues.push({
          type: "INVALID_RANGE",
          nodeId: node.id,
          message: `${node.name} has an invalid output range.`
        });
      }
    }

    const outputOwners = new Map<string, string>();
    for (const node of config.formulas) {
      for (const cell of expandRange(node.output.range)) {
        const key = `${node.output.sheet}!${cell}`;
        const owner = outputOwners.get(key);
        if (owner && owner !== node.id) {
          issues.push({
            type: "OVERLAPPING_OUTPUT",
            nodeId: node.id,
            relatedNodeIds: [owner, node.id],
            message: `Output overlap at ${key} between ${owner} and ${node.id}.`
          });
        } else {
          outputOwners.set(key, node.id);
        }
      }
    }

    const uniqueOrder = new Set(executionOrder);
    if (uniqueOrder.size !== config.formulas.length) {
      issues.push({
        type: "DEPENDENCY_CYCLE",
        message: "Formula node dependency cycle detected. Execution order was partially resolved."
      });
    }

    return issues;
  }
}

import { GraphNode, TemplateRangeMapping } from "../models/graph";

export class TemplateMappingService {
  deriveFromNodes(nodes: GraphNode[]): TemplateRangeMapping[] {
    const byId = new Map(nodes.map((node) => [node.id, node]));
    const mappings: TemplateRangeMapping[] = [];

    for (const node of nodes) {
      if (node.nodeType !== "output" || node.fileRole !== "output") {
        continue;
      }

      const formulaNodeId = node.dependencies[0];
      const formulaNode = formulaNodeId ? byId.get(formulaNodeId) : undefined;
      const source = formulaNode?.inputs[0];
      if (!source) {
        continue;
      }

      mappings.push({
        key: `${node.fileName}::${node.sheet}::${node.range}`,
        label: `${node.sheet}!${node.range}`,
        sourceRange: {
          file: source.file,
          sheet: source.sheet,
          range: source.range,
          nodeId: source.nodeId
        },
        targetRange: {
          file: node.fileName,
          sheet: node.sheet,
          range: node.range,
          nodeId: node.id
        }
      });
    }

    return mappings;
  }
}

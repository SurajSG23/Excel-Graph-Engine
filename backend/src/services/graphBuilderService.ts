import { GraphEdge, GraphNode, ParsedCell } from "../models/graph";
import { FormulaParserService } from "./formulaParserService";
import { normalizeCellAddress, toNodeId } from "../utils/cellUtils";

export class GraphBuilderService {
  constructor(private readonly formulaParserService: FormulaParserService) {}

  buildFromCells(workbookId: string, cells: ParsedCell[], sheets: string[]): { nodes: GraphNode[]; edges: GraphEdge[]; sheets: string[]; workbookId: string } {
    const nodes = cells.map((cell) => {
      const id = toNodeId(cell.sheet, cell.cell);
      return {
        id,
        sheet: cell.sheet,
        cell: normalizeCellAddress(cell.cell),
        formula: cell.formula,
        value: cell.value,
        dependencies: this.formulaParserService.extractDependencies(cell.formula, cell.sheet)
      } satisfies GraphNode;
    });

    const edges = this.buildEdges(nodes);
    return {
      workbookId,
      nodes,
      edges,
      sheets
    };
  }

  rebuildFromNodes(workbookId: string, nodes: GraphNode[], sheets: string[]): { nodes: GraphNode[]; edges: GraphEdge[]; sheets: string[]; workbookId: string } {
    const rebuiltNodes = nodes.map((node) => ({
      ...node,
      dependencies: this.formulaParserService.extractDependencies(node.formula, node.sheet)
    }));

    return {
      workbookId,
      nodes: rebuiltNodes,
      edges: this.buildEdges(rebuiltNodes),
      sheets
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
}

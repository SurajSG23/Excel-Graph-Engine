export interface GraphNode {
  id: string;
  sheet: string;
  cell: string;
  formula?: string;
  value?: number;
  dependencies: string[];
}

export interface GraphEdge {
  source: string;
  target: string;
}

export interface ValidationIssue {
  type: "CIRCULAR_DEPENDENCY" | "INVALID_FORMULA" | "MISSING_REFERENCE";
  nodeId?: string;
  message: string;
  relatedNodeIds?: string[];
}

export interface WorkbookGraph {
  workbookId: string;
  nodes: GraphNode[];
  edges: GraphEdge[];
  sheets: string[];
  validationIssues: ValidationIssue[];
  version: number;
}

export interface NodeUpdate {
  id: string;
  formula?: string;
  value?: number;
}

export interface ParsedCell {
  sheet: string;
  cell: string;
  formula?: string;
  value?: number;
}

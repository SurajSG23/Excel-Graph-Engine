export type CellValue = string | number | boolean;

export interface GraphNode {
  id: string;
  sheet: string;
  cell: string;
  formula?: string;
  value?: CellValue;
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
  value?: CellValue;
}

export interface VersionItem {
  version: number;
  timestamp: string;
  label: string;
}

export interface WorkbookResponse {
  workbook: WorkbookGraph;
  versions: VersionItem[];
}

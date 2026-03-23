export type CellValue = string | number | boolean;

export type WorkbookRole = "input" | "output" | "other";

export interface CellReference {
  file: string;
  sheet: string;
  cell: string;
  external: boolean;
  original: string;
}

export interface GraphNode {
  id: string;
  fileName: string;
  fileRole: WorkbookRole;
  sheet: string;
  cell: string;
  formula?: string;
  value?: CellValue;
  dependencies: string[];
  referenceDetails: CellReference[];
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
  files: Array<{
    fileName: string;
    role: WorkbookRole;
    sheets: string[];
    uploadName: string;
  }>;
  outputFileName: string;
  validationIssues: ValidationIssue[];
  version: number;
}

export interface NodeUpdate {
  id: string;
  formula?: string;
  value?: CellValue;
}

export interface ParsedCell {
  fileName: string;
  fileRole: WorkbookRole;
  sheet: string;
  cell: string;
  formula?: string;
  value?: CellValue;
}

export interface ParsedWorkbook {
  fileName: string;
  fileRole: WorkbookRole;
  uploadName: string;
  sheets: string[];
  cells: ParsedCell[];
}

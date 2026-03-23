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

export interface VersionItem {
  version: number;
  timestamp: string;
  label: string;
}

export interface WorkbookResponse {
  workbook: WorkbookGraph;
  versions: VersionItem[];
}

export type WorkbookOperation =
  | {
      type: "ADD_CELL";
      fileName: string;
      sheet: string;
      cell: string;
      value?: CellValue;
      formula?: string;
      fileRole?: WorkbookRole;
    }
  | {
      type: "DELETE_CELLS";
      nodeIds: string[];
    }
  | {
      type: "MOVE_CELL";
      fromNodeId: string;
      toFileName: string;
      toSheet: string;
      toCell: string;
    }
  | {
      type: "INSERT_ROW";
      fileName: string;
      sheet: string;
      index: number;
      count?: number;
    }
  | {
      type: "DELETE_ROW";
      fileName: string;
      sheet: string;
      index: number;
      count?: number;
    }
  | {
      type: "INSERT_COLUMN";
      fileName: string;
      sheet: string;
      index: number;
      count?: number;
    }
  | {
      type: "DELETE_COLUMN";
      fileName: string;
      sheet: string;
      index: number;
      count?: number;
    }
  | {
      type: "ADD_SHEET";
      fileName: string;
      sheet: string;
    }
  | {
      type: "DELETE_SHEET";
      fileName: string;
      sheet: string;
    }
  | {
      type: "RENAME_SHEET";
      fileName: string;
      fromSheet: string;
      toSheet: string;
    }
  | {
      type: "COPY_PASTE";
      sourceNodeIds: string[];
      targetFileName: string;
      targetSheet: string;
      targetAnchorCell: string;
    };

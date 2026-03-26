export type CellValue = string | number | boolean;

export type WorkbookRole = "input" | "output" | "other";
export type PipelineNodeType = "input" | "formula" | "output";

export interface RangeShape {
  rows: number;
  cols: number;
  size: number;
}

export interface NodeRangeRef {
  file: string;
  sheet: string;
  range: string;
  nodeId?: string;
}

export interface TemplateRangeMapping {
  key: string;
  label: string;
  sourceRange: NodeRangeRef;
  targetRange: NodeRangeRef;
}

/**
 * RangeReference represents a dependency on a range of cells.
 * This is the primary reference type used in the range-based pipeline model.
 */
export interface RangeReference {
  file: string;
  sheet: string;
  range: string; // e.g., "A1:A10" or "B2" (single cell treated as 1x1 range)
  external: boolean;
  original: string;
}

/**
 * @deprecated Use RangeReference instead. CellReference is retained for internal
 * formula parsing where individual cell tracking is needed during execution.
 */
export interface CellReference {
  file: string;
  sheet: string;
  cell: string;
  external: boolean;
  original: string;
}

/**
 * GraphNode represents a node in the range-based pipeline graph.
 * Each node operates on ranges (not individual cells).
 *
 * Key principle: Range is the atomic unit. All operations treat
 * ranges as the smallest addressable unit in the graph.
 */
export interface GraphNode {
  id: string;
  type?: PipelineNodeType;
  nodeType: PipelineNodeType;
  fileName: string;
  fileRole: WorkbookRole;
  sheet: string;
  range: string; // The primary identifier - always a range (e.g., "A1:A10" or "B2")
  shape: RangeShape;
  operation?: string;
  inputs: NodeRangeRef[];
  inputRanges?: NodeRangeRef[];
  output?: NodeRangeRef;
  outputRange?: NodeRangeRef;
  rangeValues?: CellValue[];
  values?: CellValue[];
  formulaTemplate?: string;

  /**
   * @internal Used only during formula execution to track per-cell formulas
   * when cells in a range have different (but pattern-similar) formulas.
   * This is an implementation detail, not part of the range-based model.
   */
  formulaByCell?: Record<string, string>;

  /**
   * @deprecated Legacy field - use `range` instead. The anchor cell can be
   * derived from the range's start position. Retained temporarily for
   * backward compatibility with mutation workflows.
   */
  cell: string;
  formula?: string;
  value?: CellValue;
  dependencies: string[]; // Node IDs (range-based), not cell IDs

  /**
   * @internal Detailed cell-level references used during execution.
   * The graph structure uses node-level dependencies; this is for execution only.
   */
  referenceDetails: CellReference[];
}

export interface GraphEdge {
  source: string;
  target: string;
}

export interface ValidationIssue {
  type:
    | "CIRCULAR_DEPENDENCY"
    | "INVALID_FORMULA"
    | "MISSING_REFERENCE"
    | "INVALID_RANGE"
    | "MISMATCHED_RANGE_SHAPE"
    | "OVERLAPPING_WRITES";
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
  templateMappings?: TemplateRangeMapping[];
  validationIssues: ValidationIssue[];
  version: number;
}

export interface NodeUpdate {
  id: string;
  formula?: string;
  value?: CellValue;
}

/**
 * ParsedCell represents raw cell data from Excel parsing.
 * This is an internal representation used during the parsing phase.
 * Cells are grouped into ranges during graph building.
 * @internal
 */
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

/**
 * WorkbookOperation defines mutations that can be applied to the graph.
 *
 * NOTE: While the graph model is range-based, some operations (like ADD_CELL)
 * still operate at cell granularity for user-facing editing. These cells
 * are grouped into ranges during subsequent graph rebuilding.
 *
 * TODO: Consider migrating to range-based operations (ADD_RANGE, DELETE_RANGE)
 * for better alignment with the pipeline model.
 */
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

export interface SpreadsheetMutationResult {
  nodes: GraphNode[];
  files: WorkbookGraph["files"];
  changedNodeIds: string[];
}

export type CellValue = string | number | boolean;

export interface PipelineRange {
  sheet: string;
  range: string;
}

export interface FormulaNodeConfig {
  id: string;
  name: string;
  inputs: PipelineRange[];
  output: PipelineRange;
  formula: string;
  structureKey: string;
  anchorCell: string;
  outputCells: string[];
}

export interface PipelineConfig {
  input: {
    id: "input";
    name: "Input";
    filePath: string;
    sheets: string[];
    ranges: PipelineRange[];
  };
  formulas: FormulaNodeConfig[];
  output: {
    id: "output";
    name: "Output";
    targetFilePath: string;
    sheets: string[];
    ranges: PipelineRange[];
  };
}

export interface PipelineGraphNode {
  id: string;
  type: "input" | "formula" | "output";
  label: string;
}

export interface PipelineGraphEdge {
  source: string;
  target: string;
}

export interface ValidationIssue {
  type: "INVALID_FORMULA" | "INVALID_RANGE" | "OVERLAPPING_OUTPUT" | "DEPENDENCY_CYCLE";
  nodeId?: string;
  message: string;
  relatedNodeIds?: string[];
}

export interface PipelineWorkbook {
  workbookId: string;
  version: number;
  config: PipelineConfig;
  graph: {
    nodes: PipelineGraphNode[];
    edges: PipelineGraphEdge[];
  };
  validationIssues: ValidationIssue[];
  executionOrder: string[];
  nodeResults: Record<string, CellValue[]>;
}

export interface PipelineNodeUpdate {
  id: string;
  formula?: string;
  inputs?: PipelineRange[];
  output?: PipelineRange;
}

export interface VersionItem {
  version: number;
  timestamp: string;
  label: string;
}

export interface WorkbookResponse {
  workbook: PipelineWorkbook;
  versions: VersionItem[];
}

export type CellValue = string | number | boolean;

export interface PipelineRange {
  sheet: string;
  range: string;
}

export interface InputNodeConfig {
  id: "input";
  name: "Input";
  filePath: string;
  sheets: string[];
  ranges: PipelineRange[];
}

export interface FormulaNodeConfig {
  id: string;
  name: string;
  inputs: PipelineRange[];
  inputTemplate: string;
  inputMappingByCell: Record<string, string[]>;
  output: PipelineRange;
  formula: string;
  formulaTemplate: string;
  formulaByCell: Record<string, string>;
  structureKey: string;
  anchorCell: string;
  outputCells: string[];
}

export interface FormulaCellEdit {
  outputCell: string;
  formula?: string;
  newOutputCell?: string;
}

export interface OutputNodeConfig {
  id: "output";
  name: "Output";
  targetFilePath: string;
  sheets: string[];
  ranges: PipelineRange[];
}

export interface PipelineConfig {
  input: InputNodeConfig;
  formulas: FormulaNodeConfig[];
  output: OutputNodeConfig;
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

export interface PipelineGraph {
  nodes: PipelineGraphNode[];
  edges: PipelineGraphEdge[];
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
  graph: PipelineGraph;
  validationIssues: ValidationIssue[];
  executionOrder: string[];
  nodeResults: Record<string, CellValue[]>;
  inputValuesByCell: Record<string, CellValue>;
}

export interface PipelineNodeUpdate {
  id: string;
  formula?: string;
  inputs?: PipelineRange[];
  inputMappingByCell?: Record<string, string[]>;
  output?: PipelineRange;
  formulaByCell?: Record<string, string>;
  cellEdits?: FormulaCellEdit[];
  inputValuesByCell?: Record<string, CellValue>;
  ranges?: PipelineRange[];
  sheets?: string[];
  filePath?: string;
  targetFilePath?: string;
}

export interface ParsedFormulaCell {
  sheet: string;
  cell: string;
  formula: string;
}

export interface ParsedWorkbookData {
  sourceFilePath: string;
  sourceFileName: string;
  targetFilePath: string;
  sheetNames: string[];
  targetSheetNames: string[];
  formulaCells: ParsedFormulaCell[];
  values: Record<string, Record<string, CellValue>>;
}

export interface VersionItem {
  version: number;
  timestamp: string;
  label: string;
}

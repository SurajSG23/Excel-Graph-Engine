import { ChangeEvent, FormEvent, useEffect, useMemo, useState } from "react";
import { UploadCard } from "./UploadCard";
import { useWorkbookStore } from "../store/workbookStore";
import { PipelineRange } from "../types/workbook";

interface SidebarProps {
  isOpen: boolean;
  onToggle: () => void;
  onResizeStart: () => void;
}

interface OutputFlowEditDraft {
  formula: string;
  inputsText: string;
  outputSheet: string;
  outputRange: string;
}

interface DetailedFormulaDraft {
  sourceOutputCell: string;
  outputCell: string;
  formula: string;
  inputsText: string;
}

interface InputCellDraft {
  id: string;
  sheet: string;
  cell: string;
}

function rangesToText(ranges: PipelineRange[]): string {
  return ranges.map((item) => `${item.sheet}!${item.range}`).join("\n");
}

function textToRanges(value: string): PipelineRange[] {
  return value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [sheet, range] = line.includes("!")
        ? (line.split("!") as [string, string])
        : (["Sheet1", line] as [string, string]);
      return { sheet: sheet.trim(), range: range.trim().toUpperCase() };
    });
}

function uniqueSheetsFromRanges(ranges: PipelineRange[]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const item of ranges) {
    const sheet = item.sheet.trim();
    if (!sheet || seen.has(sheet)) {
      continue;
    }
    seen.add(sheet);
    out.push(sheet);
  }
  return out;
}

function displayWorkbookName(pathValue: string): string {
  const normalized = pathValue.replace(/\\/g, "/");
  const fileName = normalized.split("/").pop() ?? pathValue;
  return fileName.replace(/^\d+-/, "");
}

function columnToNumber(col: string): number {
  let total = 0;
  for (const ch of col.toUpperCase()) {
    total = total * 26 + (ch.charCodeAt(0) - 64);
  }
  return total;
}

function numberToColumn(value: number): string {
  let n = value;
  let col = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

function parseCellAddress(cell: string): { col: number; row: number } | null {
  const match = cell.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
  if (!match) {
    return null;
  }
  return {
    col: columnToNumber(match[1]),
    row: Number(match[2])
  };
}

function estimateRangeCellCount(rangeText: string): number {
  const ranges = rangeText
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  let total = 0;
  for (const segment of ranges) {
    const [left, right] = segment.includes(":")
      ? (segment.split(":") as [string, string])
      : ([segment, segment] as [string, string]);
    const leftMatch = left.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
    const rightMatch = right.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
    if (!leftMatch || !rightMatch) {
      continue;
    }

    const leftCol = columnToNumber(leftMatch[1]);
    const rightCol = columnToNumber(rightMatch[1]);
    const leftRow = Number(leftMatch[2]);
    const rightRow = Number(rightMatch[2]);
    const colCount = Math.abs(rightCol - leftCol) + 1;
    const rowCount = Math.abs(rightRow - leftRow) + 1;
    total += colCount * rowCount;
  }

  return total;
}

function expandRangeCells(rangeText: string): string[] {
  const ranges = rangeText
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  const out: string[] = [];
  for (const segment of ranges) {
    const [left, right] = segment.includes(":")
      ? (segment.split(":") as [string, string])
      : ([segment, segment] as [string, string]);
    const leftMatch = left.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
    const rightMatch = right.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
    if (!leftMatch || !rightMatch) {
      continue;
    }

    const leftCol = columnToNumber(leftMatch[1]);
    const rightCol = columnToNumber(rightMatch[1]);
    const leftRow = Number(leftMatch[2]);
    const rightRow = Number(rightMatch[2]);

    const minCol = Math.min(leftCol, rightCol);
    const maxCol = Math.max(leftCol, rightCol);
    const minRow = Math.min(leftRow, rightRow);
    const maxRow = Math.max(leftRow, rightRow);

    for (let row = minRow; row <= maxRow; row += 1) {
      for (let col = minCol; col <= maxCol; col += 1) {
        out.push(`${numberToColumn(col)}${row}`);
      }
    }
  }

  return out;
}

function buildCellPreview(range: string, limit = 6): string[] {
  return expandRangeCells(range).slice(0, limit);
}

function translateFormulaForOutput(formula: string, anchorCell: string, outputCell: string, outputSheet: string): string {
  const anchor = parseCellAddress(anchorCell);
  const current = parseCellAddress(outputCell);
  if (!anchor || !current) {
    return formula;
  }

  const dRow = current.row - anchor.row;
  const dCol = current.col - anchor.col;
  return formula.replace(
    /((?:'[^']+'|[A-Za-z0-9_\.]+)!|)(\$?)([A-Z]{1,3})(\$?)([0-9]+)/g,
    (raw, sheetPrefix: string, absCol: string, colText: string, absRow: string, rowText: string) => {
      const base = parseCellAddress(`${colText}${rowText}`);
      if (!base) {
        return raw;
      }

      const nextCol = absCol ? base.col : base.col + dCol;
      const nextRow = absRow ? base.row : base.row + dRow;
      if (nextCol < 1 || nextRow < 1) {
        return raw;
      }

      const nextCell = `${numberToColumn(nextCol)}${nextRow}`;
      return `${sheetPrefix || `${outputSheet}!`}${nextCell}`;
    }
  );
}

function rangeLooksValid(value: string): boolean {
  const trimmed = value.trim().toUpperCase();
  if (!trimmed) {
    return false;
  }

  const segments = trimmed.split(",").map((item) => item.trim()).filter(Boolean);
  if (segments.length === 0) {
    return false;
  }

  return segments.every((segment) => {
    const [left, right] = segment.includes(":")
      ? (segment.split(":") as [string, string])
      : ([segment, segment] as [string, string]);
    return /^[A-Z]{1,3}[0-9]+$/.test(left) && /^[A-Z]{1,3}[0-9]+$/.test(right);
  });
}

function cellLooksValid(value: string): boolean {
  return /^[A-Z]{1,3}[0-9]+$/i.test(value.trim());
}

function rangesToDetailedCells(ranges: PipelineRange[]): InputCellDraft[] {
  const rows: InputCellDraft[] = [];
  let index = 1;

  for (const range of ranges) {
    const cells = expandRangeCells(range.range);
    for (const cell of cells) {
      rows.push({
        id: `${range.sheet}:${cell}:${index}`,
        sheet: range.sheet,
        cell
      });
      index += 1;
    }
  }

  return rows;
}

function detailedCellsToRanges(cells: InputCellDraft[]): PipelineRange[] {
  const bySheet = new Map<string, Set<string>>();

  for (const row of cells) {
    const sheet = row.sheet.trim();
    const cell = row.cell.trim().toUpperCase();
    if (!sheet || !cellLooksValid(cell)) {
      continue;
    }
    if (!bySheet.has(sheet)) {
      bySheet.set(sheet, new Set());
    }
    bySheet.get(sheet)!.add(cell);
  }

  return [...bySheet.entries()].map(([sheet, refs]) => ({
    sheet,
    range: [...refs].join(",")
  }));
}

function parseFormulaInputs(formula: string, currentSheet: string): string[] {
  const refs: string[] = [];
  const regex = /((?:'[^']+'|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g;

  for (const match of formula.matchAll(regex)) {
    const token = match[0].replace(/\$/g, "");
    if (token.includes("!")) {
      refs.push(token.toUpperCase());
    } else {
      refs.push(`${currentSheet}!${token.toUpperCase()}`);
    }
  }

  return [...new Set(refs)];
}

function normalizeInputToken(value: string, fallbackSheet: string): string | null {
  const trimmed = value.trim().replace(/\$/g, "");
  if (!trimmed) {
    return null;
  }

  if (trimmed.includes("!")) {
    const [sheet, cell] = trimmed.split("!");
    if (!sheet || !/^[A-Z]{1,3}[0-9]+$/i.test(cell)) {
      return null;
    }
    return `${sheet.trim()}!${cell.toUpperCase()}`;
  }

  if (!/^[A-Z]{1,3}[0-9]+$/i.test(trimmed)) {
    return null;
  }
  return `${fallbackSheet}!${trimmed.toUpperCase()}`;
}

function rewriteFormulaInputs(formula: string, nextInputs: string[], fallbackSheet: string): string {
  const orderedUniqueOld: string[] = [];
  const oldToNew = new Map<string, string>();
  const regex = /((?:'[^']+'|[A-Za-z0-9_\.]+)!\$?[A-Z]{1,3}\$?[0-9]+|\$?[A-Z]{1,3}\$?[0-9]+)/g;

  for (const match of formula.matchAll(regex)) {
    const token = match[0];
    if (!orderedUniqueOld.includes(token)) {
      orderedUniqueOld.push(token);
    }
  }

  orderedUniqueOld.forEach((token, index) => {
    const next = nextInputs[index];
    if (!next) {
      return;
    }

    const normalized = normalizeInputToken(next, fallbackSheet);
    if (!normalized) {
      return;
    }

    oldToNew.set(token, normalized);
  });

  return formula.replace(regex, (token) => oldToNew.get(token) ?? token);
}

export function Sidebar({ isOpen, onToggle, onResizeStart }: SidebarProps) {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedNodeId = useWorkbookStore((s) => s.selectedNodeId);
  const loading = useWorkbookStore((s) => s.loading);
  const error = useWorkbookStore((s) => s.error);
  const undo = useWorkbookStore((s) => s.undo);
  const redo = useWorkbookStore((s) => s.redo);
  const triggerExport = useWorkbookStore((s) => s.triggerExport);
  const updateFormulaNode = useWorkbookStore((s) => s.updateFormulaNode);
  const uploadFiles = useWorkbookStore((s) => s.uploadFiles);

  const formulaNode = useMemo(
    () => workbook?.config.formulas.find((node) => node.id === selectedNodeId) ?? null,
    [selectedNodeId, workbook]
  );

  const selectableSheets = useMemo(() => {
    if (!workbook) {
      return [] as string[];
    }

    const seen = new Set<string>();
    const ordered: string[] = [];

    for (const sheet of workbook.config.output.sheets) {
      if (!seen.has(sheet)) {
        seen.add(sheet);
        ordered.push(sheet);
      }
    }

    for (const sheet of workbook.config.input.sheets) {
      if (!seen.has(sheet)) {
        seen.add(sheet);
        ordered.push(sheet);
      }
    }

    for (const node of workbook.config.formulas) {
      if (!seen.has(node.output.sheet)) {
        seen.add(node.output.sheet);
        ordered.push(node.output.sheet);
      }
    }

    return ordered;
  }, [workbook]);

  const [formStatus, setFormStatus] = useState<{ type: "success" | "error"; message: string } | null>(null);

  const [inputRangesDraft, setInputRangesDraft] = useState<PipelineRange[]>([]);
  const [inputViewMode, setInputViewMode] = useState<"grouped" | "detailed">("grouped");
  const [inputDetailedDrafts, setInputDetailedDrafts] = useState<InputCellDraft[]>([]);
  const [inputStatus, setInputStatus] = useState<{ type: "success" | "error"; message: string } | null>(null);

  const [formulaText, setFormulaText] = useState("");
  const [formulaInputsText, setFormulaInputsText] = useState("");
  const [formulaOutputSheet, setFormulaOutputSheet] = useState("");
  const [formulaOutputRange, setFormulaOutputRange] = useState("");
  const [formulaViewMode, setFormulaViewMode] = useState<"grouped" | "detailed">("grouped");
  const [detailedDrafts, setDetailedDrafts] = useState<Record<string, DetailedFormulaDraft>>({});

  const [outputStatus, setOutputStatus] = useState<{ type: "success" | "error"; message: string } | null>(null);
  const [outputFlowDrafts, setOutputFlowDrafts] = useState<Record<string, OutputFlowEditDraft>>({});

  useEffect(() => {
    if (!workbook) {
      setInputRangesDraft([]);
      setInputDetailedDrafts([]);
      return;
    }
    setInputRangesDraft(workbook.config.input.ranges);
    setInputDetailedDrafts(rangesToDetailedCells(workbook.config.input.ranges));
  }, [workbook]);

  useEffect(() => {
    if (!formulaNode) {
      setFormulaText("");
      setFormulaInputsText("");
      setFormulaOutputSheet("");
      setFormulaOutputRange("");
      setDetailedDrafts({});
      return;
    }

    setFormulaText(formulaNode.formula);
    setFormulaInputsText(rangesToText(formulaNode.inputs));
    setFormulaOutputSheet(formulaNode.output.sheet);
    setFormulaOutputRange(formulaNode.output.range);

    const outputCells = formulaNode.outputCells.length > 0
      ? formulaNode.outputCells
      : expandRangeCells(formulaNode.output.range);
    const drafts: Record<string, DetailedFormulaDraft> = {};
    for (const cell of outputCells) {
      const expression = formulaNode.formulaByCell?.[cell]
        ?? translateFormulaForOutput(formulaNode.formulaTemplate || formulaNode.formula, formulaNode.anchorCell, cell, formulaNode.output.sheet);
      drafts[cell] = {
        sourceOutputCell: cell,
        outputCell: cell,
        formula: expression,
        inputsText: parseFormulaInputs(expression, formulaNode.output.sheet).join(", ")
      };
    }
    setDetailedDrafts(drafts);
  }, [formulaNode]);

  useEffect(() => {
    if (!workbook) {
      setOutputFlowDrafts({});
      return;
    }

    const drafts: Record<string, OutputFlowEditDraft> = {};
    for (const node of workbook.config.formulas) {
      drafts[node.id] = {
        formula: node.formula,
        inputsText: rangesToText(node.inputs),
        outputSheet: node.output.sheet,
        outputRange: node.output.range
      };
    }
    setOutputFlowDrafts(drafts);
  }, [workbook]);

  const inputSummary = useMemo(() => {
    const totalRanges = inputRangesDraft.length;
    const totalCells = inputRangesDraft.reduce((sum, item) => sum + estimateRangeCellCount(item.range), 0);
    return {
      totalRanges,
      totalCells,
      usedByFormulaCount: workbook?.config.formulas.length ?? 0
    };
  }, [inputRangesDraft, workbook]);

  const inputDetailedBaselineKeys = useMemo(() => {
    if (!workbook) {
      return new Set<string>();
    }
    return new Set(
      rangesToDetailedCells(workbook.config.input.ranges)
        .map((item) => `${item.sheet}!${item.cell}`)
    );
  }, [workbook]);

  const inputDetailedCurrentKeys = useMemo(
    () => new Set(inputDetailedDrafts.map((item) => `${item.sheet.trim()}!${item.cell.trim().toUpperCase()}`)),
    [inputDetailedDrafts]
  );

  const inputDetailedModifiedCount = useMemo(() => {
    let modified = 0;
    for (const key of inputDetailedCurrentKeys) {
      if (!inputDetailedBaselineKeys.has(key)) {
        modified += 1;
      }
    }
    for (const key of inputDetailedBaselineKeys) {
      if (!inputDetailedCurrentKeys.has(key)) {
        modified += 1;
      }
    }
    return modified;
  }, [inputDetailedBaselineKeys, inputDetailedCurrentKeys]);

  const inputDetailedSummary = useMemo(() => {
    const ranges = detailedCellsToRanges(inputDetailedDrafts);
    return {
      totalRanges: ranges.length,
      totalCells: inputDetailedDrafts.filter((item) => item.sheet.trim() && cellLooksValid(item.cell)).length
    };
  }, [inputDetailedDrafts]);

  const formulaOutputCells = useMemo(() => {
    if (!formulaNode) {
      return [] as string[];
    }

    if (formulaNode.outputCells.length > 0) {
      return formulaNode.outputCells;
    }

    return expandRangeCells(formulaNode.output.range);
  }, [formulaNode]);

  const formulaPreview = useMemo(() => {
    if (!formulaNode || !workbook) {
      return [] as Array<{ outputCell: string; result: string; expression: string }>;
    }

    const values = workbook.nodeResults[formulaNode.id] ?? [];
    return formulaOutputCells.slice(0, 12).map((cell, index) => {
      const translated = translateFormulaForOutput(formulaNode.formula, formulaNode.anchorCell, cell, formulaNode.output.sheet);
      return {
        outputCell: `${formulaNode.output.sheet}!${cell}`,
        result: values[index] === undefined ? "-" : String(values[index]),
        expression: translated.startsWith("=") ? translated.slice(1) : translated
      };
    });
  }, [formulaNode, formulaOutputCells, workbook]);

  const formulaDetailedRows = useMemo(() => {
    if (!formulaNode) {
      return [] as Array<{ sourceOutputCell: string; outputCell: string; formula: string; inputs: string[]; result: string; modified: boolean }>;
    }

    const values = workbook?.nodeResults[formulaNode.id] ?? [];
    const rows: Array<{ sourceOutputCell: string; outputCell: string; formula: string; inputs: string[]; result: string; modified: boolean }> = [];

    formulaOutputCells.forEach((cell, index) => {
      const defaultFormula = formulaNode.formulaByCell?.[cell]
        ?? translateFormulaForOutput(formulaNode.formulaTemplate || formulaNode.formula, formulaNode.anchorCell, cell, formulaNode.output.sheet);
      const draft = detailedDrafts[cell];
      const outputCell = draft?.outputCell || cell;
      const formula = draft?.formula || defaultFormula;
      const inputs = (draft?.inputsText || parseFormulaInputs(defaultFormula, formulaNode.output.sheet).join(", "))
        .split(",")
        .map((item) => item.trim())
        .filter(Boolean);

      const modified = Boolean(
        draft && (
          draft.outputCell !== cell ||
          draft.formula.trim() !== defaultFormula.trim() ||
          draft.inputsText.trim() !== parseFormulaInputs(defaultFormula, formulaNode.output.sheet).join(", ")
        )
      );

      rows.push({
        sourceOutputCell: cell,
        outputCell,
        formula,
        inputs,
        result: values[index] === undefined ? "-" : String(values[index]),
        modified
      });
    });

    return rows;
  }, [detailedDrafts, formulaNode, formulaOutputCells, workbook]);

  const detailedModifiedCount = useMemo(
    () => formulaDetailedRows.filter((row) => row.modified).length,
    [formulaDetailedRows]
  );

  const formulaValidationError = useMemo(() => {
    const trimmedFormula = formulaText.trim();
    if (!trimmedFormula.startsWith("=")) {
      return "Calculation must start with '='.";
    }

    const parsedInputs = textToRanges(formulaInputsText);
    if (parsedInputs.length === 0) {
      return "Add at least one input cell or range.";
    }

    if (parsedInputs.some((item) => !item.sheet || !rangeLooksValid(item.range))) {
      return "Some input cells/ranges are invalid.";
    }

    if (!formulaOutputSheet.trim() || !rangeLooksValid(formulaOutputRange)) {
      return "Choose a valid output sheet and cell/range.";
    }

    return null;
  }, [formulaInputsText, formulaOutputRange, formulaOutputSheet, formulaText]);

  const handleReplaceInputWorkbook = async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    await uploadFiles({ inputFile: file });
    setInputStatus({ type: "success", message: "Source file updated." });
    event.target.value = "";
  };

  const handleReplaceOutputWorkbook = async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    await uploadFiles({ outputFile: file });
    setOutputStatus({ type: "success", message: "Target file updated." });
    event.target.value = "";
  };

  const applyInputChanges = async (): Promise<void> => {
    if (!workbook) {
      return;
    }

    const nextRanges = inputViewMode === "grouped"
      ? inputRangesDraft.map((item) => ({ sheet: item.sheet.trim(), range: item.range.trim().toUpperCase() }))
      : detailedCellsToRanges(inputDetailedDrafts);

    if (nextRanges.length === 0) {
      setInputStatus({ type: "error", message: inputViewMode === "grouped" ? "Add at least one input range." : "Add at least one input cell." });
      return;
    }

    if (inputViewMode === "grouped") {
      if (nextRanges.some((item) => !item.sheet.trim() || !rangeLooksValid(item.range))) {
        setInputStatus({ type: "error", message: "One or more input ranges are invalid." });
        return;
      }
    } else if (inputDetailedDrafts.some((item) => !item.sheet.trim() || !cellLooksValid(item.cell))) {
      setInputStatus({ type: "error", message: "One or more detailed input cells are invalid." });
      return;
    }

    const ok = await updateFormulaNode(
      {
        id: "input",
        ranges: nextRanges,
        sheets: uniqueSheetsFromRanges(nextRanges)
      },
      inputViewMode === "grouped" ? "Edit input data" : "Edit input cells"
    );

    if (!ok) {
      setInputStatus({ type: "error", message: useWorkbookStore.getState().error ?? "Could not update input settings." });
      return;
    }

    setInputStatus({ type: "success", message: "Input settings applied and flow recalculated." });
  };

  const onFormulaSubmit = async (event: FormEvent): Promise<void> => {
    event.preventDefault();
    if (!formulaNode || formulaValidationError) {
      setFormStatus({ type: "error", message: formulaValidationError ?? "Select a calculation node first." });
      return;
    }

    const ok = await updateFormulaNode(
      {
        id: formulaNode.id,
        formula: formulaText.trim(),
        inputs: textToRanges(formulaInputsText),
        output: {
          sheet: formulaOutputSheet.trim(),
          range: formulaOutputRange.trim().toUpperCase()
        }
      },
      "Edit calculation flow"
    );

    if (!ok) {
      setFormStatus({ type: "error", message: useWorkbookStore.getState().error ?? "Could not update calculation." });
      return;
    }

    setFormStatus({ type: "success", message: "Calculation updated. Graph and results refreshed." });
  };

  const applyDetailedFormulaChanges = async (): Promise<void> => {
    if (!formulaNode) {
      return;
    }

    const edits = formulaDetailedRows
      .filter((row) => row.modified)
      .map((row) => {
        const normalizedInputs = row.inputs
          .map((item) => normalizeInputToken(item, formulaNode.output.sheet))
          .filter((item): item is string => Boolean(item));

        let nextFormula = row.formula.trim();
        nextFormula = rewriteFormulaInputs(nextFormula, normalizedInputs, formulaNode.output.sheet);

        return {
          outputCell: row.sourceOutputCell,
          formula: nextFormula,
          newOutputCell: row.outputCell.trim().toUpperCase()
        };
      });

    if (edits.length === 0) {
      setFormStatus({ type: "error", message: "No detailed changes to apply." });
      return;
    }

    if (edits.some((item) => !item.formula.startsWith("="))) {
      setFormStatus({ type: "error", message: "Each detailed formula must start with '='." });
      return;
    }

    if (edits.some((item) => !/^[A-Z]{1,3}[0-9]+$/.test(item.newOutputCell))) {
      setFormStatus({ type: "error", message: "Each output cell must be a valid cell reference like F5." });
      return;
    }

    const ok = await updateFormulaNode(
      {
        id: formulaNode.id,
        cellEdits: edits
      },
      "Edit detailed calculation cells"
    );

    if (!ok) {
      setFormStatus({ type: "error", message: useWorkbookStore.getState().error ?? "Could not apply detailed changes." });
      return;
    }

    setFormStatus({ type: "success", message: "Detailed cell changes applied and grouped logic was updated." });
  };

  const applyOutputFlowChange = async (formulaId: string): Promise<void> => {
    const draft = outputFlowDrafts[formulaId];
    if (!draft) {
      return;
    }

    const inputs = textToRanges(draft.inputsText);
    if (!draft.formula.trim().startsWith("=")) {
      setOutputStatus({ type: "error", message: "Calculation must start with '='." });
      return;
    }

    if (inputs.length === 0 || inputs.some((item) => !item.sheet || !rangeLooksValid(item.range))) {
      setOutputStatus({ type: "error", message: "Please provide valid input cells/ranges." });
      return;
    }

    if (!draft.outputSheet.trim() || !rangeLooksValid(draft.outputRange)) {
      setOutputStatus({ type: "error", message: "Please provide a valid output sheet and cell/range." });
      return;
    }

    const ok = await updateFormulaNode(
      {
        id: formulaId,
        formula: draft.formula.trim(),
        inputs,
        output: {
          sheet: draft.outputSheet.trim(),
          range: draft.outputRange.trim().toUpperCase()
        }
      },
      "Edit output flow"
    );

    if (!ok) {
      setOutputStatus({ type: "error", message: useWorkbookStore.getState().error ?? "Could not update output flow." });
      return;
    }

    setOutputStatus({ type: "success", message: "Output flow updated and recalculated." });
  };

  return (
    <aside className={`dashboard-sidebar ${isOpen ? "is-open" : "is-closed"}`}>
      <div className="sidebar-header">
        <h1>Excel Pipeline Engine</h1>
        <button
          type="button"
          className="sidebar-toggle icon-button"
          onClick={onToggle}
          aria-label={isOpen ? "Collapse controls panel" : "Expand controls panel"}
        >
          <span className={isOpen ? "icon-chevron-left" : "icon-chevron-right"} aria-hidden="true" />
        </button>
      </div>

      <div className="sidebar-content">
        <UploadCard />

        <section className="panel sidebar-actions-panel">
          <h3>Pipeline Actions</h3>
          <div className="sidebar-actions-grid">
            <button type="button" className="sidebar-action-btn" onClick={() => void undo()} disabled={loading || !workbook}>Undo</button>
            <button type="button" className="sidebar-action-btn" onClick={() => void redo()} disabled={loading || !workbook}>Redo</button>
            <button type="button" className="sidebar-action-btn sidebar-action-btn-primary" onClick={() => void triggerExport()} disabled={loading || !workbook}>Export</button>
          </div>
        </section>

        {error && <div className="error-banner">{error}</div>}
        {loading && <div className="loading-banner">Recomputing flow...</div>}

        <section className="panel inspector">
          <h3>Node Details</h3>
          <p className="inspector-subtitle">Build logic as a simple flow: Inputs {"->"} Calculation {"->"} Output.</p>

          {!workbook && <p className="node-empty-state">Upload a workbook to start editing your flow.</p>}

          {workbook && selectedNodeId === "input" && (
            <div className="flow-panel">
              <h4>Inputs</h4>
              <p className="node-details-lead">Choose the source file and the exact cells/ranges to read.</p>

              <div className="view-toggle-row" role="tablist" aria-label="Input view mode">
                <button
                  type="button"
                  className={`view-toggle-btn ${inputViewMode === "grouped" ? "is-active" : ""}`}
                  onClick={() => setInputViewMode("grouped")}
                >
                  Grouped View
                </button>
                <button
                  type="button"
                  className={`view-toggle-btn ${inputViewMode === "detailed" ? "is-active" : ""}`}
                  onClick={() => setInputViewMode("detailed")}
                >
                  Detailed View
                </button>
              </div>

              <div className="flow-block">
                <div className="flow-block-title-row">
                  <h5>Source File</h5>
                  <span>{displayWorkbookName(workbook.config.input.filePath)}</span>
                </div>
                <label className="file-swap-control">
                  <span>Change source file</span>
                  <input type="file" accept=".xlsx" onChange={(event) => void handleReplaceInputWorkbook(event)} disabled={loading} />
                </label>
              </div>

              {inputViewMode === "grouped" && (
                <div className="flow-block">
                  <div className="flow-block-title-row">
                    <h5>Sheets and Ranges</h5>
                    <button
                      type="button"
                      className="ghost-btn"
                      onClick={() =>
                        setInputRangesDraft((prev) => [
                          ...prev,
                          { sheet: workbook.config.input.sheets[0] ?? "Sheet1", range: "A1:A10" }
                        ])
                      }
                      disabled={loading}
                    >
                      Add Range
                    </button>
                  </div>

                  <div className="flow-card-list">
                    {inputRangesDraft.map((item, index) => {
                      const preview = buildCellPreview(item.range);
                      return (
                        <article className="flow-card" key={`input-range:${index}`}>
                          <div className="flow-edit-grid">
                            <label>
                              <span>Sheet</span>
                              <select
                                value={item.sheet}
                                onChange={(event) =>
                                  setInputRangesDraft((prev) =>
                                    prev.map((row, rowIndex) =>
                                      rowIndex === index ? { ...row, sheet: event.target.value } : row
                                    )
                                  )
                                }
                                disabled={loading}
                              >
                                {selectableSheets.map((sheet) => (
                                  <option key={`input-sheet:${sheet}`} value={sheet}>{sheet}</option>
                                ))}
                              </select>
                            </label>

                            <label>
                              <span>Range</span>
                              <input
                                value={item.range}
                                onChange={(event) =>
                                  setInputRangesDraft((prev) =>
                                    prev.map((row, rowIndex) =>
                                      rowIndex === index ? { ...row, range: event.target.value.toUpperCase() } : row
                                    )
                                  )
                                }
                                placeholder="A1:A100"
                                disabled={loading}
                              />
                            </label>
                          </div>

                          <div className="mini-preview">
                            <span>Preview</span>
                            <div className="mini-preview-grid">
                              {preview.length === 0 && <em>No preview</em>}
                              {preview.map((cell) => (
                                <code key={`preview:${index}:${cell}`}>{cell}</code>
                              ))}
                            </div>
                          </div>

                          <button
                            type="button"
                            className="danger-text-btn"
                            onClick={() =>
                              setInputRangesDraft((prev) => prev.filter((_, rowIndex) => rowIndex !== index))
                            }
                            disabled={loading}
                          >
                            Remove
                          </button>
                        </article>
                      );
                    })}
                  </div>
                </div>
              )}

              {inputViewMode === "detailed" && (
                <div className="flow-block">
                  <div className="flow-block-title-row">
                    <h5>Detailed Input Cells</h5>
                    <span>{inputDetailedModifiedCount} modified</span>
                  </div>
                  <div className="node-result-table-wrap">
                    <table className="node-result-table detailed-flow-table">
                      <thead>
                        <tr>
                          <th>Sheet</th>
                          <th>Cell</th>
                          <th>Used By</th>
                          <th>Action</th>
                        </tr>
                      </thead>
                      <tbody>
                        {inputDetailedDrafts.map((row) => {
                          const key = `${row.sheet.trim()}!${row.cell.trim().toUpperCase()}`;
                          const isModified = !inputDetailedBaselineKeys.has(key);
                          return (
                            <tr key={`input-cell:${row.id}`} className={isModified ? "row-modified" : ""}>
                              <td>
                                <select
                                  className="table-cell-input"
                                  value={row.sheet}
                                  onChange={(event) =>
                                    setInputDetailedDrafts((prev) =>
                                      prev.map((item) => item.id === row.id ? { ...item, sheet: event.target.value } : item)
                                    )
                                  }
                                  disabled={loading}
                                >
                                  {selectableSheets.map((sheet) => (
                                    <option key={`input-detail-sheet:${row.id}:${sheet}`} value={sheet}>{sheet}</option>
                                  ))}
                                </select>
                              </td>
                              <td>
                                <input
                                  className="table-cell-input"
                                  value={row.cell}
                                  onChange={(event) =>
                                    setInputDetailedDrafts((prev) =>
                                      prev.map((item) => item.id === row.id ? { ...item, cell: event.target.value.toUpperCase() } : item)
                                    )
                                  }
                                  disabled={loading}
                                />
                              </td>
                              <td>
                                {workbook.config.formulas.filter((node) =>
                                  node.inputs.some((item) => item.sheet === row.sheet && expandRangeCells(item.range).includes(row.cell.toUpperCase()))
                                ).length}
                              </td>
                              <td>
                                <button
                                  type="button"
                                  className="danger-text-btn"
                                  onClick={() => setInputDetailedDrafts((prev) => prev.filter((item) => item.id !== row.id))}
                                  disabled={loading}
                                >
                                  Remove
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  <button
                    type="button"
                    className="ghost-btn"
                    onClick={() =>
                      setInputDetailedDrafts((prev) => [
                        ...prev,
                        {
                          id: `new:${prev.length + 1}`,
                          sheet: selectableSheets[0] ?? "Sheet1",
                          cell: "A1"
                        }
                      ])
                    }
                    disabled={loading}
                  >
                    Add Cell
                  </button>
                </div>
              )}

              <div className="flow-block">
                <h5>Summary</h5>
                <div className="flow-stat-grid">
                  <div><span>Cells Read</span><strong>{inputViewMode === "grouped" ? inputSummary.totalCells : inputDetailedSummary.totalCells}</strong></div>
                  <div><span>Ranges</span><strong>{inputViewMode === "grouped" ? inputSummary.totalRanges : inputDetailedSummary.totalRanges}</strong></div>
                  <div><span>Used By</span><strong>{inputSummary.usedByFormulaCount} calculations</strong></div>
                </div>
              </div>

              {inputStatus && (
                <div className={`form-status-banner form-status-${inputStatus.type}`}>{inputStatus.message}</div>
              )}
              <button type="button" onClick={() => void applyInputChanges()} disabled={loading}>Apply Input Changes</button>
            </div>
          )}

          {formulaNode && (
            <div className="flow-panel">
              <h4>Calculation</h4>
              <p className="node-details-lead">Edit how inputs transform into outputs.</p>

              <div className="view-toggle-row" role="tablist" aria-label="Formula view mode">
                <button
                  type="button"
                  className={`view-toggle-btn ${formulaViewMode === "grouped" ? "is-active" : ""}`}
                  onClick={() => setFormulaViewMode("grouped")}
                >
                  Grouped View
                </button>
                <button
                  type="button"
                  className={`view-toggle-btn ${formulaViewMode === "detailed" ? "is-active" : ""}`}
                  onClick={() => setFormulaViewMode("detailed")}
                >
                  Detailed View
                </button>
              </div>

              <div className="flow-block">
                <h5>Flow</h5>
                <div className="flow-line">
                  <span>{formulaNode.inputs.map((item) => `${item.sheet}!${item.range}`).join(", ") || "Inputs"}</span>
                  <span className="flow-arrow">{"->"}</span>
                  <span>{formulaNode.formulaTemplate || formulaNode.formula}</span>
                  <span className="flow-arrow">{"->"}</span>
                  <span>{formulaNode.output.sheet}!{formulaNode.output.range}</span>
                </div>
              </div>

              {formulaViewMode === "grouped" && (
                <form className="formula-form" onSubmit={(event) => void onFormulaSubmit(event)}>
                  <div className="flow-block">
                    <h5>Inputs</h5>
                    <textarea
                      rows={4}
                      value={formulaInputsText}
                      onChange={(event) => setFormulaInputsText(event.target.value)}
                      placeholder="Sheet1!A1\nSheet1!B2"
                    />
                  </div>

                  <div className="flow-block">
                    <h5>Calculation</h5>
                    <textarea
                      rows={3}
                      value={formulaText}
                      onChange={(event) => setFormulaText(event.target.value)}
                      placeholder="=A1+B2"
                    />
                  </div>

                  <div className="flow-block">
                    <h5>Output</h5>
                    <div className="flow-edit-grid">
                      <label>
                        <span>Sheet</span>
                        <select value={formulaOutputSheet} onChange={(event) => setFormulaOutputSheet(event.target.value)}>
                          {selectableSheets.length === 0 && <option value="">No sheets</option>}
                          {selectableSheets.map((sheetName) => (
                            <option key={`formula-sheet:${sheetName}`} value={sheetName}>{sheetName}</option>
                          ))}
                        </select>
                      </label>

                      <label>
                        <span>Cell / Range</span>
                        <input
                          value={formulaOutputRange}
                          onChange={(event) => setFormulaOutputRange(event.target.value.toUpperCase())}
                          placeholder="E2 or E2:E20"
                        />
                      </label>
                    </div>
                  </div>

                  {formulaValidationError && <div className="error-banner">{formulaValidationError}</div>}
                  {formStatus && <div className={`form-status-banner form-status-${formStatus.type}`}>{formStatus.message}</div>}
                  <button type="submit" disabled={loading || Boolean(formulaValidationError)}>Apply Grouped Changes</button>
                </form>
              )}

              {formulaViewMode === "detailed" && (
                <div className="flow-block">
                  <div className="flow-block-title-row">
                    <h5>Detailed Cell Flow</h5>
                    <span>{detailedModifiedCount} modified</span>
                  </div>
                  <div className="node-result-table-wrap">
                    <table className="node-result-table detailed-flow-table">
                      <thead>
                        <tr>
                          <th>Output Cell</th>
                          <th>Formula</th>
                          <th>Inputs</th>
                          <th>Result</th>
                        </tr>
                      </thead>
                      <tbody>
                        {formulaDetailedRows.map((row) => (
                          <tr key={`${formulaNode.id}:detail:${row.sourceOutputCell}`} className={row.modified ? "row-modified" : ""}>
                            <td>
                              <input
                                className="table-cell-input"
                                value={row.outputCell}
                                onChange={(event) =>
                                  setDetailedDrafts((prev) => ({
                                    ...prev,
                                    [row.sourceOutputCell]: {
                                      sourceOutputCell: row.sourceOutputCell,
                                      outputCell: event.target.value.toUpperCase(),
                                      formula: prev[row.sourceOutputCell]?.formula ?? row.formula,
                                      inputsText: prev[row.sourceOutputCell]?.inputsText ?? row.inputs.join(", ")
                                    }
                                  }))
                                }
                                disabled={loading}
                              />
                            </td>
                            <td>
                              <input
                                className="table-cell-input"
                                value={row.formula}
                                onChange={(event) =>
                                  setDetailedDrafts((prev) => ({
                                    ...prev,
                                    [row.sourceOutputCell]: {
                                      sourceOutputCell: row.sourceOutputCell,
                                      outputCell: prev[row.sourceOutputCell]?.outputCell ?? row.outputCell,
                                      formula: event.target.value,
                                      inputsText: prev[row.sourceOutputCell]?.inputsText ?? row.inputs.join(", ")
                                    }
                                  }))
                                }
                                disabled={loading}
                              />
                            </td>
                            <td>
                              <input
                                className="table-cell-input"
                                value={row.inputs.join(", ")}
                                onChange={(event) =>
                                  setDetailedDrafts((prev) => ({
                                    ...prev,
                                    [row.sourceOutputCell]: {
                                      sourceOutputCell: row.sourceOutputCell,
                                      outputCell: prev[row.sourceOutputCell]?.outputCell ?? row.outputCell,
                                      formula: prev[row.sourceOutputCell]?.formula ?? row.formula,
                                      inputsText: event.target.value
                                    }
                                  }))
                                }
                                disabled={loading}
                              />
                            </td>
                            <td>{row.result}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {formStatus && <div className={`form-status-banner form-status-${formStatus.type}`}>{formStatus.message}</div>}
                  <button type="button" disabled={loading} onClick={() => void applyDetailedFormulaChanges()}>
                    Apply Detailed Cell Changes
                  </button>
                </div>
              )}

              <div className="flow-block">
                <h5>Preview</h5>
                <div className="node-result-table-wrap">
                  <table className="node-result-table">
                    <thead>
                      <tr>
                        <th>Output Cell</th>
                        <th>Expression</th>
                        <th>Result</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formulaPreview.map((item) => (
                        <tr key={`${formulaNode.id}:${item.outputCell}`}>
                          <td>{item.outputCell}</td>
                          <td><code>{item.expression}</code></td>
                          <td>{item.result}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {workbook && selectedNodeId === "output" && (
            <div className="flow-panel">
              <h4>Output</h4>
              <p className="node-details-lead">Review and edit final output flows in one place.</p>

              <div className="flow-block">
                <div className="flow-block-title-row">
                  <h5>Target File</h5>
                  <span>{displayWorkbookName(workbook.config.output.targetFilePath)}</span>
                </div>
                <label className="file-swap-control">
                  <span>Change target file</span>
                  <input type="file" accept=".xlsx" onChange={(event) => void handleReplaceOutputWorkbook(event)} disabled={loading} />
                </label>
              </div>

              <div className="flow-block">
                <h5>Output Flow</h5>
                <div className="flow-card-list">
                  {workbook.config.formulas.map((node) => {
                    const draft = outputFlowDrafts[node.id];
                    if (!draft) {
                      return null;
                    }

                    return (
                      <article className="flow-card" key={`output-flow:${node.id}`}>
                        <div className="flow-line compact">
                          <span>{node.inputs.map((item) => `${item.sheet}!${item.range}`).join(", ") || "Inputs"}</span>
                          <span className="flow-arrow">{"->"}</span>
                          <span>{node.formula}</span>
                          <span className="flow-arrow">{"->"}</span>
                          <span>{node.output.sheet}!{node.output.range}</span>
                        </div>

                        <label>
                          <span>Inputs</span>
                          <textarea
                            rows={3}
                            value={draft.inputsText}
                            onChange={(event) =>
                              setOutputFlowDrafts((prev) => ({
                                ...prev,
                                [node.id]: { ...prev[node.id], inputsText: event.target.value }
                              }))
                            }
                            disabled={loading}
                          />
                        </label>

                        <label>
                          <span>Calculation</span>
                          <input
                            value={draft.formula}
                            onChange={(event) =>
                              setOutputFlowDrafts((prev) => ({
                                ...prev,
                                [node.id]: { ...prev[node.id], formula: event.target.value }
                              }))
                            }
                            disabled={loading}
                          />
                        </label>

                        <div className="flow-edit-grid">
                          <label>
                            <span>Output Sheet</span>
                            <select
                              value={draft.outputSheet}
                              onChange={(event) =>
                                setOutputFlowDrafts((prev) => ({
                                  ...prev,
                                  [node.id]: { ...prev[node.id], outputSheet: event.target.value }
                                }))
                              }
                              disabled={loading}
                            >
                              {selectableSheets.map((sheet) => (
                                <option key={`out-sheet:${node.id}:${sheet}`} value={sheet}>{sheet}</option>
                              ))}
                            </select>
                          </label>

                          <label>
                            <span>Output Cell / Range</span>
                            <input
                              value={draft.outputRange}
                              onChange={(event) =>
                                setOutputFlowDrafts((prev) => ({
                                  ...prev,
                                  [node.id]: { ...prev[node.id], outputRange: event.target.value.toUpperCase() }
                                }))
                              }
                              disabled={loading}
                            />
                          </label>
                        </div>

                        <button type="button" onClick={() => void applyOutputFlowChange(node.id)} disabled={loading}>
                          Apply This Output Flow
                        </button>
                      </article>
                    );
                  })}
                </div>
              </div>

              <div className="flow-block">
                <h5>Lineage</h5>
                <div className="lineage-list">
                  {workbook.config.formulas.map((node) => (
                    <p key={`lineage:${node.id}`}>
                      {node.inputs.map((item) => `${item.sheet}!${item.range}`).join(" + ")} {"->"} {node.formula} {"->"} {node.output.sheet}!{node.output.range}
                    </p>
                  ))}
                </div>
              </div>

              {outputStatus && <div className={`form-status-banner form-status-${outputStatus.type}`}>{outputStatus.message}</div>}
            </div>
          )}

          {workbook && !selectedNodeId && (
            <p className="node-empty-state">Select Inputs, a Calculation node, or Output in the graph to edit.</p>
          )}
        </section>

        {workbook && (
          <section className="panel validation-panel">
            <h3>Validation</h3>
            {workbook.validationIssues.length === 0 ? (
              <p className="validation-empty">No validation issues.</p>
            ) : (
              <ul className="validation-list">
                {workbook.validationIssues.map((issue, index) => (
                  <li key={`${issue.type}:${issue.nodeId ?? "global"}:${index}`} className="validation-item">
                    <span className="validation-badge">{issue.type}</span>
                    <span>{issue.message}</span>
                  </li>
                ))}
              </ul>
            )}
          </section>
        )}
      </div>

      <button
        type="button"
        className="sidebar-resize-handle"
        onPointerDown={(event) => {
          if (event.pointerType === "mouse" || event.pointerType === "pen") {
            event.preventDefault();
          }
          onResizeStart();
        }}
        aria-label="Resize sidebar"
      />
    </aside>
  );
}

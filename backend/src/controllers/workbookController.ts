import { Request, Response } from "express";
import fs from "node:fs";
import {
  excelParser,
  executionEngine,
  exportService,
  pipelineBuilder,
  pipelineValidator,
  workbookSessionService
} from "../services/serviceContainer";
import { CellValue, FormulaNodeConfig, PipelineNodeUpdate, PipelineRange } from "../models/pipeline";
import { collapseCellsToRange, expandRange, extractFormulaRefs, parseCell, parseRefToken, toCell } from "../core/node_models";

function resolveUploadPath(req: Request, field: string): string | undefined {
  const files = (req.files as Record<string, Express.Multer.File[]>) ?? {};
  return files[field]?.[0]?.path;
}

function normalizeRanges(ranges: PipelineRange[] | undefined): PipelineRange[] | undefined {
  if (!ranges) {
    return undefined;
  }

  return ranges
    .filter((item) => typeof item?.sheet === "string" && typeof item?.range === "string")
    .map((item) => ({
      sheet: item.sheet.trim(),
      range: item.range.trim().toUpperCase()
    }));
}

function normalizeSheets(sheets: string[] | undefined): string[] | undefined {
  if (!sheets) {
    return undefined;
  }

  const out = sheets
    .map((item) => item.trim())
    .filter(Boolean);

  return [...new Set(out)];
}

function normalizeCellRef(cell: string): string {
  return cell.trim().toUpperCase();
}

function normalizeCellFormulaMap(map: Record<string, string> | undefined): Record<string, string> {
  if (!map) {
    return {};
  }

  const out: Record<string, string> = {};
  for (const [cell, formula] of Object.entries(map)) {
    const normalizedCell = normalizeCellRef(cell);
    const normalizedFormula = formula.trim();
    if (!normalizedCell || !normalizedFormula) {
      continue;
    }
    out[normalizedCell] = normalizedFormula;
  }
  return out;
}

function normalizeCellValue(value: unknown): CellValue {
  if (typeof value === "number" || typeof value === "boolean") {
    return value;
  }

  const text = String(value ?? "").trim();
  if (/^-?\d+(?:\.\d+)?$/.test(text)) {
    return Number(text);
  }
  if (/^true$/i.test(text)) {
    return true;
  }
  if (/^false$/i.test(text)) {
    return false;
  }
  return text;
}

function normalizeInputValuesMap(map: Record<string, CellValue> | undefined, defaultSheet: string): Record<string, CellValue> {
  if (!map) {
    return {};
  }

  const out: Record<string, CellValue> = {};
  for (const [key, raw] of Object.entries(map)) {
    const parsed = parseRefToken(key, defaultSheet);
    const cell = normalizeCellRef(parsed.cell);
    if (!cell) {
      continue;
    }
    out[`${parsed.sheet}!${cell}`] = normalizeCellValue(raw);
  }
  return out;
}

function cloneValues(values: Record<string, Record<string, CellValue>>): Record<string, Record<string, CellValue>> {
  const out: Record<string, Record<string, CellValue>> = {};
  for (const [sheet, entries] of Object.entries(values)) {
    out[sheet] = { ...entries };
  }
  return out;
}

function applyInputValues(
  values: Record<string, Record<string, CellValue>>,
  updates: Record<string, CellValue>
): Record<string, Record<string, CellValue>> {
  const out = cloneValues(values);
  for (const [key, value] of Object.entries(updates)) {
    const [sheet, cell] = key.split("!") as [string, string];
    if (!out[sheet]) {
      out[sheet] = {};
    }
    out[sheet][cell] = value;
  }
  return out;
}

function buildInputValueSnapshot(
  ranges: PipelineRange[],
  values: Record<string, Record<string, CellValue>>
): Record<string, CellValue> {
  const out: Record<string, CellValue> = {};
  for (const item of ranges) {
    for (const cell of expandRange(item.range)) {
      const value = values[item.sheet]?.[cell];
      out[`${item.sheet}!${cell}`] = value ?? "";
    }
  }
  return out;
}

function sortCells(cells: string[]): string[] {
  return [...cells].sort((left, right) => {
    const l = parseCell(left);
    const r = parseCell(right);
    if (!l || !r) {
      return left.localeCompare(right);
    }
    if (l.row !== r.row) {
      return l.row - r.row;
    }
    return l.col - r.col;
  });
}

function buildOutputRange(cells: string[]): string {
  return sortCells(cells).join(",");
}

function translateFormula(formula: string, fromCell: string, toCellRef: string, outputSheet: string): string {
  const from = parseCell(fromCell);
  const to = parseCell(toCellRef);
  if (!from || !to) {
    return formula;
  }

  const dRow = to.row - from.row;
  const dCol = to.col - from.col;

  return formula.replace(
    /((?:'[^']+'|[A-Za-z0-9_\.]+)!|)(\$?)([A-Z]{1,3})(\$?)([0-9]+)/g,
    (_raw, sheetPrefix: string, absCol: string, colText: string, absRow: string, rowText: string) => {
      const base = parseCell(`${colText}${rowText}`);
      if (!base) {
        return _raw;
      }

      const nextCol = absCol ? base.col : base.col + dCol;
      const nextRow = absRow ? base.row : base.row + dRow;
      if (nextCol < 1 || nextRow < 1) {
        return _raw;
      }

      const nextCell = toCell(nextRow, nextCol);
      return `${sheetPrefix || `${outputSheet}!`}${nextCell}`;
    }
  );
}

function deriveInputsFromFormulas(formulas: string[], outputSheet: string): PipelineRange[] {
  const bySheet = new Map<string, Set<string>>();

  for (const formula of formulas) {
    for (const ref of extractFormulaRefs(formula, outputSheet)) {
      if (!bySheet.has(ref.sheet)) {
        bySheet.set(ref.sheet, new Set());
      }
      bySheet.get(ref.sheet)!.add(ref.cell);
    }
  }

  return [...bySheet.entries()].map(([sheet, cells]) => ({
    sheet,
    range: collapseCellsToRange([...cells])
  }));
}

function deriveInputTemplate(mapping: string[]): string {
  return mapping
    .map((item) => {
      const [sheet, cell] = item.includes("!")
        ? (item.split("!") as [string, string])
        : (["", item] as [string, string]);
      const col = cell.replace(/[0-9]/g, "").replace(/\$/g, "").toUpperCase();
      return sheet ? `${sheet}!${col}` : col;
    })
    .join(", ");
}

function deriveInputMappingByCell(formulaByCell: Record<string, string>, outputSheet: string): Record<string, string[]> {
  const out: Record<string, string[]> = {};
  for (const [cell, formula] of Object.entries(formulaByCell)) {
    out[cell] = extractFormulaRefs(formula, outputSheet).map((ref) => `${ref.sheet}!${ref.cell}`);
  }
  return out;
}

function buildFormulaForCell(node: FormulaNodeConfig, cell: string): string {
  const normalized = normalizeCellRef(cell);
  if (node.formulaByCell?.[normalized]) {
    return node.formulaByCell[normalized];
  }
  const template = node.formulaTemplate || node.formula;
  return translateFormula(template, node.anchorCell, normalized, node.output.sheet);
}

function parseInputsForCell(formula: string, outputSheet: string): string[] {
  const refs = extractFormulaRefs(formula, outputSheet)
    .map((item) => `${item.sheet}!${item.cell}`)
    .filter(Boolean);
  return [...new Set(refs)];
}

function ensureUniqueId(existingIds: Set<string>, preferred: string): string {
  if (!existingIds.has(preferred)) {
    existingIds.add(preferred);
    return preferred;
  }

  let counter = 2;
  while (existingIds.has(`${preferred}-${counter}`)) {
    counter += 1;
  }
  const next = `${preferred}-${counter}`;
  existingIds.add(next);
  return next;
}

function applyFormulaPatch(node: FormulaNodeConfig, patch: PipelineNodeUpdate, existingIds: Set<string>): FormulaNodeConfig[] {
  const baseOutput = patch.output
    ? {
        sheet: patch.output.sheet.trim(),
        range: patch.output.range.trim().toUpperCase()
      }
    : node.output;
  const outputCells = patch.output
    ? sortCells(expandRange(baseOutput.range))
    : sortCells(node.outputCells.length > 0 ? node.outputCells : expandRange(node.output.range));

  const nextTemplate = typeof patch.formula === "string" ? patch.formula.trim() : (node.formulaTemplate || node.formula);
  const normalizedByCellPatch = normalizeCellFormulaMap(patch.formulaByCell);

  if (!patch.cellEdits || patch.cellEdits.length === 0) {
    const formulaByCell: Record<string, string> = {};
    for (const cell of outputCells) {
      formulaByCell[cell] = normalizedByCellPatch[cell] ?? translateFormula(nextTemplate, outputCells[0] ?? node.anchorCell, cell, baseOutput.sheet);
    }

    const formulas = Object.values(formulaByCell);
    const inputMappingByCell = patch.inputMappingByCell ?? deriveInputMappingByCell(formulaByCell, baseOutput.sheet);
    return [
      {
        ...node,
        formula: nextTemplate,
        formulaTemplate: nextTemplate,
        formulaByCell,
        inputMappingByCell,
        inputTemplate: deriveInputTemplate(inputMappingByCell[outputCells[0] ?? node.anchorCell] ?? []),
        inputs: normalizeRanges(patch.inputs) ?? deriveInputsFromFormulas(formulas, baseOutput.sheet),
        output: baseOutput,
        anchorCell: outputCells[0] ?? node.anchorCell,
        outputCells
      }
    ];
  }

  const editsByCell = new Map(
    patch.cellEdits
      .map((item) => ({
        ...item,
        outputCell: normalizeCellRef(item.outputCell),
        newOutputCell: item.newOutputCell ? normalizeCellRef(item.newOutputCell) : undefined,
        formula: item.formula?.trim()
      }))
      .filter((item) => item.outputCell)
      .map((item) => [item.outputCell, item])
  );

  const remainingCells: string[] = [];
  const remainingFormulaByCell: Record<string, string> = {};
  const detachedNodes: FormulaNodeConfig[] = [];

  for (const cell of outputCells) {
    const edit = editsByCell.get(cell);
    const baseFormula = buildFormulaForCell(
      {
        ...node,
        formulaTemplate: nextTemplate,
        formulaByCell: {
          ...node.formulaByCell,
          ...normalizedByCellPatch
        }
      },
      cell
    );

    if (!edit) {
      remainingCells.push(cell);
      remainingFormulaByCell[cell] = baseFormula;
      continue;
    }

    const movedCell = edit.newOutputCell ?? cell;
    const nextFormula = edit.formula
      ? edit.formula
      : (movedCell !== cell ? translateFormula(baseFormula, cell, movedCell, baseOutput.sheet) : baseFormula);
    const nextId = ensureUniqueId(existingIds, `${node.id}__${movedCell}`);

    detachedNodes.push({
      ...node,
      id: nextId,
      name: `${node.name} ${movedCell}`,
      inputs: deriveInputsFromFormulas([nextFormula], baseOutput.sheet),
      inputTemplate: deriveInputTemplate(parseInputsForCell(nextFormula, baseOutput.sheet)),
      inputMappingByCell: {
        [movedCell]: parseInputsForCell(nextFormula, baseOutput.sheet)
      },
      output: {
        sheet: baseOutput.sheet,
        range: movedCell
      },
      formula: nextFormula,
      formulaTemplate: nextFormula,
      formulaByCell: {
        [movedCell]: nextFormula
      },
      structureKey: `${node.structureKey}:${movedCell}`,
      anchorCell: movedCell,
      outputCells: [movedCell]
    });
  }

  const merged: FormulaNodeConfig[] = [];
  if (remainingCells.length > 0) {
    const remainingFormulas = Object.values(remainingFormulaByCell);
    const remainingInputMappingByCell = deriveInputMappingByCell(remainingFormulaByCell, baseOutput.sheet);
    merged.push({
      ...node,
      formula: nextTemplate,
      formulaTemplate: nextTemplate,
      formulaByCell: remainingFormulaByCell,
      inputMappingByCell: remainingInputMappingByCell,
      inputTemplate: deriveInputTemplate(remainingInputMappingByCell[sortCells(remainingCells)[0] ?? node.anchorCell] ?? []),
      inputs: deriveInputsFromFormulas(remainingFormulas, baseOutput.sheet),
      output: {
        sheet: baseOutput.sheet,
        range: buildOutputRange(remainingCells)
      },
      anchorCell: sortCells(remainingCells)[0] ?? node.anchorCell,
      outputCells: sortCells(remainingCells)
    });
  }

  return [...merged, ...detachedNodes];
}

export class WorkbookController {
  uploadWorkbook(req: Request, res: Response): void {
    try {
      const inputPath = resolveUploadPath(req, "input") ?? resolveUploadPath(req, "file");
      const outputPath = resolveUploadPath(req, "output");

      if (!inputPath && !outputPath) {
        res.status(400).json({ message: "No workbook uploaded." });
        return;
      }

      const sourcePath = inputPath ?? outputPath!;
      const targetPath = outputPath ?? sourcePath;
      const parsed = excelParser.parse(sourcePath, targetPath);
      const built = pipelineBuilder.build(parsed);
      const ordered = built.executionOrder
        .map((id) => built.config.formulas.find((item) => item.id === id))
        .filter((item): item is NonNullable<typeof item> => Boolean(item));
      const execution = executionEngine.execute(parsed, ordered);
      const validationIssues = [
        ...pipelineValidator.validate(built.config, built.executionOrder),
        ...execution.issues.map((issue) => ({
          type: "INVALID_FORMULA" as const,
          nodeId: issue.nodeId,
          message: issue.message
        }))
      ];

      const existingWorkbookId = typeof req.body?.workbookId === "string" ? req.body.workbookId : undefined;

      if (existingWorkbookId) {
        const session = workbookSessionService.getSession(existingWorkbookId);
        if (!session) {
          res.status(404).json({ message: "Workbook not found." });
          return;
        }

        const workbook = workbookSessionService.updateWorkbook(
          existingWorkbookId,
          {
            workbookId: existingWorkbookId,
            config: built.config,
            graph: built.graph,
            validationIssues,
            executionOrder: built.executionOrder,
            nodeResults: execution.nodeResults,
            inputValuesByCell: buildInputValueSnapshot(built.config.input.ranges, execution.values)
          },
          "Upload workbook"
        );
        workbookSessionService.setParsedWorkbook(existingWorkbookId, {
          ...parsed,
          values: execution.values
        });

        res.status(200).json({
          workbook,
          versions: workbookSessionService.getVersions(existingWorkbookId)
        });
        return;
      }

      const workbook = workbookSessionService.createSession(
        {
          config: built.config,
          graph: built.graph,
          validationIssues,
          executionOrder: built.executionOrder,
          nodeResults: execution.nodeResults,
          inputValuesByCell: buildInputValueSnapshot(built.config.input.ranges, execution.values)
        },
        {
          ...parsed,
          values: execution.values
        }
      );

      res.status(200).json({
        workbook,
        versions: workbookSessionService.getVersions(workbook.workbookId)
      });
    } catch (error) {
      res.status(500).json({
        message: "Failed to build pipeline from workbook.",
        detail: error instanceof Error ? error.message : "Unknown error"
      });
    }
  }

  recomputeWorkbook(req: Request, res: Response): void {
    try {
      const { workbookId, updates, label } = req.body as {
        workbookId?: string;
        updates?: PipelineNodeUpdate[];
        label?: string;
      };

      if (!workbookId) {
        res.status(400).json({ message: "workbookId is required." });
        return;
      }

      const session = workbookSessionService.getSession(workbookId);
      if (!session) {
        res.status(404).json({ message: "Workbook not found." });
        return;
      }

      const safeUpdates = updates ?? [];
      const inputPatch = safeUpdates.find((item) => item.id === "input");
      const outputPatch = safeUpdates.find((item) => item.id === "output");
      const normalizedInputValues = normalizeInputValuesMap(inputPatch?.inputValuesByCell, session.workbook.config.input.sheets[0] ?? "Sheet1");

      const existingIds = new Set(session.workbook.config.formulas.map((item) => item.id));
      const nextFormulas = session.workbook.config.formulas.flatMap((node) => {
        const patch = safeUpdates.find((item) => item.id === node.id);
        if (!patch) {
          return [node];
        }

        existingIds.delete(node.id);
        return applyFormulaPatch(node, patch, existingIds);
      });

      const nextInput = {
        ...session.workbook.config.input,
        filePath:
          typeof inputPatch?.filePath === "string" && inputPatch.filePath.trim().length > 0
            ? inputPatch.filePath.trim()
            : session.workbook.config.input.filePath,
        sheets: normalizeSheets(inputPatch?.sheets) ?? session.workbook.config.input.sheets,
        ranges: normalizeRanges(inputPatch?.ranges) ?? session.workbook.config.input.ranges
      };

      const nextOutputRanges = normalizeRanges(outputPatch?.ranges);
      const nextOutput = {
        ...session.workbook.config.output,
        targetFilePath:
          typeof outputPatch?.targetFilePath === "string" && outputPatch.targetFilePath.trim().length > 0
            ? outputPatch.targetFilePath.trim()
            : session.workbook.config.output.targetFilePath,
        sheets: normalizeSheets(outputPatch?.sheets) ?? session.workbook.config.output.sheets,
        ranges: nextOutputRanges ?? nextFormulas.map((item) => item.output)
      };

      const rebuilt = pipelineBuilder.rebuild({
        input: nextInput,
        formulas: nextFormulas,
        output: nextOutput
      });

      const ordered = rebuilt.executionOrder
        .map((id) => rebuilt.config.formulas.find((item) => item.id === id))
        .filter((item): item is NonNullable<typeof item> => Boolean(item));
      const startingValues = applyInputValues(session.parsedWorkbook.values, normalizedInputValues);
      const execution = executionEngine.execute(session.parsedWorkbook, ordered, startingValues);
      const validationIssues = [
        ...pipelineValidator.validate(rebuilt.config, rebuilt.executionOrder),
        ...execution.issues.map((issue) => ({
          type: "INVALID_FORMULA" as const,
          nodeId: issue.nodeId,
          message: issue.message
        }))
      ];

      const workbook = workbookSessionService.updateWorkbook(
        workbookId,
        {
          workbookId,
          config: rebuilt.config,
          graph: rebuilt.graph,
          validationIssues,
          executionOrder: rebuilt.executionOrder,
          nodeResults: execution.nodeResults,
          inputValuesByCell: buildInputValueSnapshot(rebuilt.config.input.ranges, execution.values)
        },
        label ?? "Edit formula node"
      );

      workbookSessionService.setParsedWorkbook(workbookId, {
        ...session.parsedWorkbook,
        values: execution.values
      });

      res.status(200).json({
        workbook,
        versions: workbookSessionService.getVersions(workbookId)
      });
    } catch (error) {
      res.status(500).json({
        message: "Failed to recompute pipeline.",
        detail: error instanceof Error ? error.message : "Unknown error"
      });
    }
  }

  undo(req: Request, res: Response): void {
    try {
      const { workbookId } = req.body as { workbookId?: string };
      if (!workbookId) {
        res.status(400).json({ message: "workbookId is required." });
        return;
      }

      const workbook = workbookSessionService.undo(workbookId);
      res.status(200).json({
        workbook,
        versions: workbookSessionService.getVersions(workbookId)
      });
    } catch (error) {
      res.status(400).json({
        message: "No earlier versions found.",
        detail: error instanceof Error ? error.message : "Unknown error"
      });
    }
  }

  redo(req: Request, res: Response): void {
    try {
      const { workbookId } = req.body as { workbookId?: string };
      if (!workbookId) {
        res.status(400).json({ message: "workbookId is required." });
        return;
      }

      const workbook = workbookSessionService.redo(workbookId);
      res.status(200).json({
        workbook,
        versions: workbookSessionService.getVersions(workbookId)
      });
    } catch (error) {
      res.status(400).json({
        message: "You are on the latest version.",
        detail: error instanceof Error ? error.message : "Unknown error"
      });
    }
  }

  exportWorkbook(req: Request, res: Response): void {
    try {
      const { workbookId } = req.body as { workbookId?: string };
      if (!workbookId) {
        res.status(400).json({ message: "workbookId is required." });
        return;
      }

      const session = workbookSessionService.getSession(workbookId);
      if (!session) {
        res.status(404).json({ message: "Workbook not found." });
        return;
      }

      if (!fs.existsSync(session.parsedWorkbook.targetFilePath)) {
        res.status(400).json({ message: "Target workbook file is missing on disk." });
        return;
      }

      const exportPath = exportService.exportWorkbook(session.workbook, session.parsedWorkbook);
      res.download(exportPath, `pipeline-${workbookId}.xlsx`);
    } catch (error) {
      res.status(500).json({
        message: "Failed to export workbook.",
        detail: error instanceof Error ? error.message : "Unknown error"
      });
    }
  }

  runPipeline(req: Request, res: Response): void {
    try {
      const { workbookId, label } = req.body as { workbookId?: string; label?: string };
      if (!workbookId) {
        res.status(400).json({ message: "workbookId is required." });
        return;
      }

      const session = workbookSessionService.getSession(workbookId);
      if (!session) {
        res.status(404).json({ message: "Workbook not found." });
        return;
      }

      const rebuilt = pipelineBuilder.rebuild(session.workbook.config);
      const ordered = rebuilt.executionOrder
        .map((id) => rebuilt.config.formulas.find((item) => item.id === id))
        .filter((item): item is NonNullable<typeof item> => Boolean(item));
      const execution = executionEngine.execute(session.parsedWorkbook, ordered, session.parsedWorkbook.values);
      const validationIssues = [
        ...pipelineValidator.validate(rebuilt.config, rebuilt.executionOrder),
        ...execution.issues.map((issue) => ({
          type: "INVALID_FORMULA" as const,
          nodeId: issue.nodeId,
          message: issue.message
        }))
      ];

      const workbook = workbookSessionService.updateWorkbook(
        workbookId,
        {
          workbookId,
          config: rebuilt.config,
          graph: rebuilt.graph,
          validationIssues,
          executionOrder: rebuilt.executionOrder,
          nodeResults: execution.nodeResults,
          inputValuesByCell: buildInputValueSnapshot(rebuilt.config.input.ranges, execution.values)
        },
        label ?? "Run pipeline"
      );

      workbookSessionService.setParsedWorkbook(workbookId, {
        ...session.parsedWorkbook,
        values: execution.values
      });

      res.status(200).json({
        workbook,
        versions: workbookSessionService.getVersions(workbookId)
      });
    } catch (error) {
      res.status(500).json({
        message: "Failed to run pipeline.",
        detail: error instanceof Error ? error.message : "Unknown error"
      });
    }
  }
}

export const workbookController = new WorkbookController();

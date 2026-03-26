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
import { PipelineNodeUpdate, PipelineRange } from "../models/pipeline";

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
      const validationIssues = pipelineValidator.validate(built.config, built.executionOrder);

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
            nodeResults: execution.nodeResults
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
          nodeResults: execution.nodeResults
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

      const nextFormulas = session.workbook.config.formulas.map((node) => {
        const patch = updates?.find((item) => item.id === node.id);
        if (!patch) {
          return node;
        }

        const formula = typeof patch.formula === "string" ? patch.formula.trim() : node.formula;
        const inputs = normalizeRanges(patch.inputs) ?? node.inputs;
        const output = patch.output
          ? {
              sheet: patch.output.sheet.trim(),
              range: patch.output.range.trim().toUpperCase()
            }
          : node.output;

        return {
          ...node,
          formula,
          inputs,
          output,
          outputCells: output.range
            .split(",")
            .map((segment) => segment.trim())
            .filter(Boolean)
            .flatMap((segment) => {
              const [left, right] = segment.includes(":")
                ? (segment.split(":") as [string, string])
                : ([segment, segment] as [string, string]);
              const [lCol, lRow] = [left.replace(/[0-9]/g, ""), Number(left.replace(/[A-Z]/g, ""))];
              const [rCol, rRow] = [right.replace(/[0-9]/g, ""), Number(right.replace(/[A-Z]/g, ""))];
              if (!lCol || !rCol || !Number.isFinite(lRow) || !Number.isFinite(rRow)) {
                return [] as string[];
              }
              const colToNum = (col: string): number => {
                let total = 0;
                for (const ch of col) {
                  total = total * 26 + (ch.charCodeAt(0) - 64);
                }
                return total;
              };
              const numToCol = (num: number): string => {
                let n = num;
                let out = "";
                while (n > 0) {
                  const rem = (n - 1) % 26;
                  out = String.fromCharCode(65 + rem) + out;
                  n = Math.floor((n - 1) / 26);
                }
                return out;
              };
              const minCol = Math.min(colToNum(lCol), colToNum(rCol));
              const maxCol = Math.max(colToNum(lCol), colToNum(rCol));
              const minRow = Math.min(lRow, rRow);
              const maxRow = Math.max(lRow, rRow);
              const cells: string[] = [];
              for (let row = minRow; row <= maxRow; row += 1) {
                for (let col = minCol; col <= maxCol; col += 1) {
                  cells.push(`${numToCol(col)}${row}`);
                }
              }
              return cells;
            })
        };
      });

      const rebuilt = pipelineBuilder.rebuild({
        ...session.workbook.config,
        formulas: nextFormulas,
        output: {
          ...session.workbook.config.output,
          ranges: nextFormulas.map((item) => item.output)
        }
      });

      const ordered = rebuilt.executionOrder
        .map((id) => rebuilt.config.formulas.find((item) => item.id === id))
        .filter((item): item is NonNullable<typeof item> => Boolean(item));
      const execution = executionEngine.execute(session.parsedWorkbook, ordered, session.parsedWorkbook.values);
      const validationIssues = pipelineValidator.validate(rebuilt.config, rebuilt.executionOrder);

      const workbook = workbookSessionService.updateWorkbook(
        workbookId,
        {
          workbookId,
          config: rebuilt.config,
          graph: rebuilt.graph,
          validationIssues,
          executionOrder: rebuilt.executionOrder,
          nodeResults: execution.nodeResults
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

      const computed = executionEngineService.recompute(session.workbook.nodes);
      const validationIssues = validationService.validate(computed.nodes, session.workbook.files);

      const updatedWorkbook = workbookSessionService.updateWorkbook(
        workbookId,
        {
          ...session.workbook,
          workbookId,
          nodes: computed.nodes,
          templateMappings: templateMappingService.deriveFromNodes(computed.nodes),
          validationIssues: [...validationIssues, ...computed.issues]
        },
        label || "Run pipeline"
      );

      res.status(200).json({
        workbook: updatedWorkbook,
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

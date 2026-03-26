import { Request, Response } from "express";
import {
  excelParserService,
  executionEngineService,
  exportService,
  fileRegistryService,
  graphBuilderService,
  templateMappingService,
  validationService,
  workbookMutationService,
  workbookSessionService
} from "../services/serviceContainer";
import { NodeUpdate, ParsedWorkbook, WorkbookOperation, WorkbookRole } from "../models/graph";

interface UploadItem {
  path: string;
  originalname: string;
  role: WorkbookRole;
}

function toRole(value: unknown): WorkbookRole | null {
  if (typeof value !== "string") {
    return null;
  }

  const normalized = value.trim().toLowerCase();
  if (normalized === "input" || normalized === "output") {
    return normalized;
  }
  return null;
}

function collectUploads(req: Request): UploadItem[] {
  const files = (req.files as Record<string, Express.Multer.File[]>) ?? {};
  const entries: UploadItem[] = [];

  for (const item of files.input ?? []) {
    entries.push({
      path: item.path,
      originalname: item.originalname,
      role: "input"
    });
  }

  for (const item of files.output ?? []) {
    entries.push({
      path: item.path,
      originalname: item.originalname,
      role: "output"
    });
  }

  for (const item of files.file ?? []) {
    const role = toRole(req.body?.role);
    entries.push({
      path: item.path,
      originalname: item.originalname,
      role: role ?? "other"
    });
  }

  return entries;
}

function resolveOutputFileName(parsed: ParsedWorkbook[], fallback?: string): string {
  return parsed.find((item) => item.fileRole === "output")?.fileName ?? fallback ?? parsed[0]?.fileName ?? "";
}

export class WorkbookController {
  uploadWorkbook(req: Request, res: Response): void {
    try {
      const uploadItems = collectUploads(req);
      if (uploadItems.length === 0) {
        res.status(400).json({ message: "No file uploaded. Provide input/output files or a file with role=input|output." });
        return;
      }

      for (const item of uploadItems) {
        if (item.role === "other") {
          res.status(400).json({ message: "Labeled upload must include role=input or role=output." });
          return;
        }
      }

      const parsedIncoming = uploadItems.map((item) =>
        excelParserService.parseWorkbook(item.path, item.originalname, item.role)
      );

      const existingWorkbookId = typeof req.body?.workbookId === "string" ? req.body.workbookId : undefined;

      if (existingWorkbookId) {
        const session = workbookSessionService.getSession(existingWorkbookId);
        if (!session) {
          res.status(404).json({ message: "Workbook not found." });
          return;
        }

        const mergedParsed = fileRegistryService.upsertFiles(existingWorkbookId, parsedIncoming);
        const rebuilt = graphBuilderService.buildFromWorkbooks(mergedParsed);
        const validationIssues = validationService.validate(rebuilt.nodes, rebuilt.files);
        const computed = executionEngineService.recompute(rebuilt.nodes);
        const outputFileName = resolveOutputFileName(mergedParsed, session.workbook.outputFileName);

        const updatedWorkbook = workbookSessionService.updateWorkbook(
          existingWorkbookId,
          {
            workbookId: existingWorkbookId,
            nodes: computed.nodes,
            edges: rebuilt.edges,
            sheets: rebuilt.sheets,
            files: rebuilt.files,
            outputFileName,
            templateMappings: templateMappingService.deriveFromNodes(computed.nodes),
            validationIssues: [...validationIssues, ...computed.issues]
          },
          "Upload workbook"
        );

        res.status(200).json({
          workbook: updatedWorkbook,
          versions: workbookSessionService.getVersions(existingWorkbookId)
        });
        return;
      }

      const initial = graphBuilderService.buildFromWorkbooks(parsedIncoming);
      const validationIssues = validationService.validate(initial.nodes, initial.files);
      const computed = executionEngineService.recompute(initial.nodes);
      const outputFileName = resolveOutputFileName(parsedIncoming);

      const created = workbookSessionService.createSession({
        nodes: computed.nodes,
        edges: initial.edges,
        sheets: initial.sheets,
        files: initial.files,
        outputFileName,
        templateMappings: templateMappingService.deriveFromNodes(computed.nodes),
        validationIssues: [...validationIssues, ...computed.issues]
      });

      fileRegistryService.upsertFiles(created.workbookId, parsedIncoming);

      res.status(200).json({
        workbook: created,
        versions: workbookSessionService.getVersions(created.workbookId)
      });
    } catch (error) {
      res.status(500).json({
        message: "Failed to parse workbook.",
        detail: error instanceof Error ? error.message : "Unknown error"
      });
    }
  }

  recomputeWorkbook(req: Request, res: Response): void {
    try {
      const { workbookId, updates, label } = req.body as {
        workbookId?: string;
        updates?: NodeUpdate[];
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

      const changedNodeIds: string[] = [];
      const nodeMap = new Map(session.workbook.nodes.map((node) => [node.id, { ...node }]));

      for (const update of updates ?? []) {
        const target = nodeMap.get(update.id);
        if (!target) {
          continue;
        }

        if (typeof update.formula === "string") {
          target.formula = update.formula.trim() === "" ? undefined : update.formula;
          if (target.nodeType === "formula" && target.formulaByCell) {
            const nextFormula = target.formula;
            const nextMap: Record<string, string> = {};
            for (const key of Object.keys(target.formulaByCell)) {
              if (nextFormula) {
                nextMap[key] = nextFormula;
              }
            }
            target.formulaByCell = nextMap;
          }
          changedNodeIds.push(target.id);
        }

        if (
          typeof update.value === "number" ||
          typeof update.value === "string" ||
          typeof update.value === "boolean"
        ) {
          target.value = update.value;
          if (!changedNodeIds.includes(target.id)) {
            changedNodeIds.push(target.id);
          }
        }

        nodeMap.set(target.id, target);
      }

      const rebuilt = graphBuilderService.rebuildFromNodes([...nodeMap.values()], session.workbook.files);
      const validationIssues = validationService.validate(rebuilt.nodes, rebuilt.files);
      const computed = executionEngineService.recompute(rebuilt.nodes, changedNodeIds);

      const updatedWorkbook = workbookSessionService.updateWorkbook(
        workbookId,
        {
          workbookId,
          ...rebuilt,
          outputFileName: session.workbook.outputFileName,
          templateMappings: templateMappingService.deriveFromNodes(computed.nodes),
          validationIssues: [...validationIssues, ...computed.issues],
          nodes: computed.nodes
        },
        label || "Formula edit"
      );

      res.status(200).json({
        workbook: updatedWorkbook,
        versions: workbookSessionService.getVersions(workbookId)
      });
    } catch (error) {
      res.status(500).json({
        message: "Failed to recompute workbook.",
        detail: error instanceof Error ? error.message : "Unknown error"
      });
    }
  }

  applyOperations(req: Request, res: Response): void {
    try {
      const { workbookId, operations, label } = req.body as {
        workbookId?: string;
        operations?: WorkbookOperation[];
        label?: string;
      };

      if (!workbookId) {
        res.status(400).json({ message: "workbookId is required." });
        return;
      }

      if (!Array.isArray(operations) || operations.length === 0) {
        res.status(400).json({ message: "operations array is required." });
        return;
      }

      const session = workbookSessionService.getSession(workbookId);
      if (!session) {
        res.status(404).json({ message: "Workbook not found." });
        return;
      }

      const mutated = workbookMutationService.applyOperations(session.workbook, operations);
      const rebuilt = graphBuilderService.rebuildFromNodes(mutated.nodes, mutated.files);
      const validationIssues = validationService.validate(rebuilt.nodes, rebuilt.files);
      const computed = executionEngineService.recompute(rebuilt.nodes, mutated.changedNodeIds);

      const updatedWorkbook = workbookSessionService.updateWorkbook(
        workbookId,
        {
          workbookId,
          ...rebuilt,
          outputFileName: session.workbook.outputFileName,
          templateMappings: templateMappingService.deriveFromNodes(computed.nodes),
          validationIssues: [...validationIssues, ...computed.issues],
          nodes: computed.nodes
        },
        label || "Spreadsheet operation"
      );

      res.status(200).json({
        workbook: updatedWorkbook,
        versions: workbookSessionService.getVersions(workbookId)
      });
    } catch (error) {
      res.status(500).json({
        message: "Failed to apply operations.",
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
        message: "You're on the latest version.",
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

      const exportPath = exportService.exportWorkbook(
        session.workbook.nodes,
        workbookId,
        session.workbook.outputFileName
      );

      res.download(exportPath, `excel-graph-engine-${workbookId}.xlsx`);
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

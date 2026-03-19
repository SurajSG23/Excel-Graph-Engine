import { Request, Response } from "express";
import {
  excelParserService,
  executionEngineService,
  exportService,
  graphBuilderService,
  validationService,
  workbookSessionService
} from "../services/serviceContainer";
import { NodeUpdate } from "../models/graph";

export class WorkbookController {
  uploadWorkbook(req: Request, res: Response): void {
    try {
      if (!req.file?.path) {
        res.status(400).json({ message: "No file uploaded." });
        return;
      }

      const parsed = excelParserService.parseWorkbook(req.file.path);
      const initial = graphBuilderService.buildFromCells("pending", parsed.cells, parsed.sheets);
      const validationIssues = validationService.validate(initial.nodes);
      const computed = executionEngineService.recompute(initial.nodes);

      const created = workbookSessionService.createSession({
        nodes: computed.nodes,
        edges: initial.edges,
        sheets: parsed.sheets,
        validationIssues: [...validationIssues, ...computed.issues]
      });

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
          changedNodeIds.push(target.id);
        }

        if (typeof update.value === "number") {
          target.value = update.value;
          if (!changedNodeIds.includes(target.id)) {
            changedNodeIds.push(target.id);
          }
        }

        nodeMap.set(target.id, target);
      }

      const rebuilt = graphBuilderService.rebuildFromNodes(workbookId, [...nodeMap.values()], session.workbook.sheets);
      const validationIssues = validationService.validate(rebuilt.nodes);
      const computed = executionEngineService.recompute(rebuilt.nodes, changedNodeIds);

      const updatedWorkbook = workbookSessionService.updateWorkbook(
        workbookId,
        {
          ...rebuilt,
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
        session.workbook.sheets,
        workbookId
      );

      res.download(exportPath, `excel-graph-engine-${workbookId}.xlsx`);
    } catch (error) {
      res.status(500).json({
        message: "Failed to export workbook.",
        detail: error instanceof Error ? error.message : "Unknown error"
      });
    }
  }
}

export const workbookController = new WorkbookController();

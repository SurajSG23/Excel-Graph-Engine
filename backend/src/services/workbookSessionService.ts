import { v4 as uuidv4 } from "uuid";
import { ParsedWorkbookData, PipelineWorkbook, VersionItem } from "../models/pipeline";

interface WorkbookSession {
  workbook: PipelineWorkbook;
  parsedWorkbook: ParsedWorkbookData;
  versions: VersionItem[];
  undoStack: PipelineWorkbook[];
  redoStack: PipelineWorkbook[];
}

export class WorkbookSessionService {
  private readonly sessions = new Map<string, WorkbookSession>();

  createSession(
    workbook: Omit<PipelineWorkbook, "workbookId" | "version">,
    parsedWorkbook: ParsedWorkbookData
  ): PipelineWorkbook {
    const workbookId = uuidv4();
    const sessionWorkbook: PipelineWorkbook = {
      ...workbook,
      workbookId,
      version: 1
    };

    this.sessions.set(workbookId, {
      workbook: sessionWorkbook,
      parsedWorkbook,
      versions: [{ version: 1, timestamp: new Date().toISOString(), label: "Initial upload" }],
      undoStack: [],
      redoStack: []
    });

    return sessionWorkbook;
  }

  getSession(workbookId: string): WorkbookSession | undefined {
    return this.sessions.get(workbookId);
  }

  updateWorkbook(
    workbookId: string,
    nextWorkbook: Omit<PipelineWorkbook, "version">,
    label = "Recompute"
  ): PipelineWorkbook {
    const session = this.sessions.get(workbookId);
    if (!session) {
      throw new Error("Workbook session not found.");
    }

    const version = session.workbook.version + 1;
    const updated = {
      ...nextWorkbook,
      workbookId,
      version
    };

    session.undoStack.push(session.workbook);
    session.redoStack = [];
    session.workbook = updated;
    session.versions.push({
      version,
      timestamp: new Date().toISOString(),
      label
    });

    this.sessions.set(workbookId, session);
    return updated;
  }

  setParsedWorkbook(workbookId: string, parsedWorkbook: ParsedWorkbookData): void {
    const session = this.sessions.get(workbookId);
    if (!session) {
      throw new Error("Workbook session not found.");
    }
    session.parsedWorkbook = parsedWorkbook;
    this.sessions.set(workbookId, session);
  }

  getVersions(workbookId: string): VersionItem[] {
    return this.sessions.get(workbookId)?.versions ?? [];
  }

  undo(workbookId: string): PipelineWorkbook {
    const session = this.sessions.get(workbookId);
    if (!session) {
      throw new Error("Workbook session not found.");
    }

    const previous = session.undoStack.pop();
    if (!previous) {
      throw new Error("Nothing to undo.");
    }

    session.redoStack.push(session.workbook);
    session.workbook = {
      ...previous,
      workbookId,
      version: session.workbook.version + 1
    };

    session.versions.push({
      version: session.workbook.version,
      timestamp: new Date().toISOString(),
      label: "Undo"
    });

    this.sessions.set(workbookId, session);
    return session.workbook;
  }

  redo(workbookId: string): PipelineWorkbook {
    const session = this.sessions.get(workbookId);
    if (!session) {
      throw new Error("Workbook session not found.");
    }

    const next = session.redoStack.pop();
    if (!next) {
      throw new Error("Nothing to redo.");
    }

    session.undoStack.push(session.workbook);
    session.workbook = {
      ...next,
      workbookId,
      version: session.workbook.version + 1
    };

    session.versions.push({
      version: session.workbook.version,
      timestamp: new Date().toISOString(),
      label: "Redo"
    });

    this.sessions.set(workbookId, session);
    return session.workbook;
  }
}

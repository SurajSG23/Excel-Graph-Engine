import { v4 as uuidv4 } from "uuid";
import { WorkbookGraph } from "../models/graph";

interface WorkbookSession {
  workbook: WorkbookGraph;
  versions: Array<{ version: number; timestamp: string; label: string }>;
}

export class WorkbookSessionService {
  private readonly sessions = new Map<string, WorkbookSession>();

  createSession(workbook: Omit<WorkbookGraph, "workbookId" | "version">): WorkbookGraph {
    const workbookId = uuidv4();
    const sessionWorkbook: WorkbookGraph = {
      ...workbook,
      workbookId,
      version: 1
    };

    this.sessions.set(workbookId, {
      workbook: sessionWorkbook,
      versions: [{ version: 1, timestamp: new Date().toISOString(), label: "Initial upload" }]
    });

    return sessionWorkbook;
  }

  getSession(workbookId: string): WorkbookSession | undefined {
    return this.sessions.get(workbookId);
  }

  updateWorkbook(
    workbookId: string,
    nextWorkbook: Omit<WorkbookGraph, "version">,
    label = "Recompute"
  ): WorkbookGraph {
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

    session.workbook = updated;
    session.versions.push({
      version,
      timestamp: new Date().toISOString(),
      label
    });

    this.sessions.set(workbookId, session);
    return updated;
  }

  getVersions(workbookId: string): Array<{ version: number; timestamp: string; label: string }> {
    return this.sessions.get(workbookId)?.versions ?? [];
  }
}

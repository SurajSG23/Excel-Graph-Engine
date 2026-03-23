import { ParsedWorkbook, WorkbookRole } from "../models/graph";

interface RegisteredFile {
  role: WorkbookRole;
  workbook: ParsedWorkbook;
}

export class FileRegistryService {
  private readonly registry = new Map<string, Map<string, RegisteredFile>>();

  getFiles(workbookId: string): ParsedWorkbook[] {
    const files = this.registry.get(workbookId);
    if (!files) {
      return [];
    }

    return [...files.values()]
      .map((entry) => entry.workbook)
      .sort((left, right) => this.roleRank(left.fileRole) - this.roleRank(right.fileRole));
  }

  upsertFiles(workbookId: string, workbooks: ParsedWorkbook[]): ParsedWorkbook[] {
    const existing = this.registry.get(workbookId) ?? new Map<string, RegisteredFile>();

    for (const workbook of workbooks) {
      const key = workbook.fileRole === "other" ? workbook.fileName : workbook.fileRole;
      existing.set(key, {
        role: workbook.fileRole,
        workbook
      });
    }

    this.registry.set(workbookId, existing);
    return this.getFiles(workbookId);
  }

  delete(workbookId: string): void {
    this.registry.delete(workbookId);
  }

  private roleRank(role: WorkbookRole): number {
    if (role === "input") return 0;
    if (role === "output") return 1;
    return 2;
  }
}

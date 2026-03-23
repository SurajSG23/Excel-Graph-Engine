import { create } from "zustand";
import {
  applyWorkbookOperations,
  exportWorkbook,
  recomputeWorkbook,
  redoWorkbook,
  undoWorkbook,
  uploadWorkbook
} from "../services/api";
import { NodeUpdate, VersionItem, WorkbookGraph, WorkbookOperation } from "../types/workbook";

interface WorkbookState {
  workbook: WorkbookGraph | null;
  versions: VersionItem[];
  selectedNodeId: string | null;
  selectedFile: string | "ALL";
  selectedSheet: string | "ALL";
  searchText: string;
  showZeroDependencyNodes: boolean;
  loading: boolean;
  error: string | null;
  uploadFiles: (payload: {
    inputFile?: File;
    outputFile?: File;
    labeledFile?: File;
    role?: "input" | "output";
  }) => Promise<void>;
  setSelectedNode: (id: string | null) => void;
  setSelectedFile: (file: string | "ALL") => void;
  setSelectedSheet: (sheet: string | "ALL") => void;
  setSearchText: (value: string) => void;
  setShowZeroDependencyNodes: (value: boolean) => void;
  applyUpdate: (updates: NodeUpdate[], label?: string) => Promise<void>;
  applyOperations: (operations: WorkbookOperation[], label?: string) => Promise<void>;
  undo: () => Promise<void>;
  redo: () => Promise<void>;
  triggerExport: () => Promise<void>;
}

export const useWorkbookStore = create<WorkbookState>((set, get) => ({
  workbook: null,
  versions: [],
  selectedNodeId: null,
  selectedFile: "ALL",
  selectedSheet: "ALL",
  searchText: "",
  showZeroDependencyNodes: true,
  loading: false,
  error: null,

  async uploadFiles(payload) {
    set({ loading: true, error: null });
    try {
      const response = await uploadWorkbook({
        workbookId: get().workbook?.workbookId,
        ...payload
      });

      const nextSelectedFile = get().selectedFile === "ALL"
        ? "ALL"
        : response.workbook.files.some((file) => file.fileName === get().selectedFile)
          ? get().selectedFile
          : "ALL";

      set({
        workbook: response.workbook,
        versions: response.versions,
        selectedNodeId: null,
        selectedFile: nextSelectedFile,
        selectedSheet: "ALL",
        loading: false
      });
    } catch (error) {
      set({
        loading: false,
        error: error instanceof Error ? error.message : "Upload failed"
      });
    }
  },

  setSelectedNode(id) {
    set({ selectedNodeId: id });
  },

  setSelectedFile(file) {
    set({ selectedFile: file, selectedSheet: "ALL" });
  },

  setSelectedSheet(sheet) {
    set({ selectedSheet: sheet });
  },

  setSearchText(value) {
    set({ searchText: value });
  },

  setShowZeroDependencyNodes(value) {
    set({ showZeroDependencyNodes: value });
  },

  async applyUpdate(updates, label) {
    const workbookId = get().workbook?.workbookId;
    if (!workbookId || updates.length === 0) {
      return;
    }

    set({ loading: true, error: null });
    try {
      const response = await recomputeWorkbook(workbookId, updates, label);
      set({
        workbook: response.workbook,
        versions: response.versions,
        loading: false
      });
    } catch (error) {
      set({
        loading: false,
        error: error instanceof Error ? error.message : "Recompute failed"
      });
    }
  },

  async triggerExport() {
    const workbook = get().workbook;
    if (!workbook) {
      return;
    }

    set({ loading: true, error: null });
    try {
      const blob = await exportWorkbook(workbook.workbookId);
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = `excel-graph-engine-${workbook.workbookId}.xlsx`;
      anchor.click();
      URL.revokeObjectURL(url);
      set({ loading: false });
    } catch (error) {
      set({
        loading: false,
        error: error instanceof Error ? error.message : "Export failed"
      });
    }
  },

  async applyOperations(operations, label) {
    const workbookId = get().workbook?.workbookId;
    if (!workbookId || operations.length === 0) {
      return;
    }

    set({ loading: true, error: null });
    try {
      const response = await applyWorkbookOperations(workbookId, operations, label);
      set({
        workbook: response.workbook,
        versions: response.versions,
        loading: false
      });
    } catch (error) {
      set({
        loading: false,
        error: error instanceof Error ? error.message : "Operation failed"
      });
    }
  },

  async undo() {
    const workbookId = get().workbook?.workbookId;
    if (!workbookId) {
      return;
    }

    set({ loading: true, error: null });
    try {
      const response = await undoWorkbook(workbookId);
      set({
        workbook: response.workbook,
        versions: response.versions,
        loading: false
      });
    } catch (error) {
      set({
        loading: false,
        error: error instanceof Error ? error.message : "Undo failed"
      });
    }
  },

  async redo() {
    const workbookId = get().workbook?.workbookId;
    if (!workbookId) {
      return;
    }

    set({ loading: true, error: null });
    try {
      const response = await redoWorkbook(workbookId);
      set({
        workbook: response.workbook,
        versions: response.versions,
        loading: false
      });
    } catch (error) {
      set({
        loading: false,
        error: error instanceof Error ? error.message : "Redo failed"
      });
    }
  }
}));

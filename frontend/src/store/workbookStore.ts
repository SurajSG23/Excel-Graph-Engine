import { create } from "zustand";
import {
  exportWorkbook,
  recomputeWorkbook,
  runPipeline as runPipelineRequest,
  redoWorkbook,
  undoWorkbook,
  uploadWorkbook,
} from "../services/api";
import {
  FormulaNodeConfig,
  PipelineNodeUpdate,
  PipelineWorkbook,
  VersionItem,
} from "../types/workbook";

interface WorkbookState {
  workbook: PipelineWorkbook | null;
  versions: VersionItem[];
  selectedNodeId: string | null;
  loading: boolean;
  error: string | null;
  uploadFiles: (payload: {
    inputFile?: File;
    outputFile?: File;
  }) => Promise<void>;
  setSelectedNode: (id: string | null) => void;
  updateFormulaNode: (
    update: PipelineNodeUpdate,
    label?: string,
  ) => Promise<boolean>;
  undo: () => Promise<void>;
  redo: () => Promise<void>;
  triggerExport: () => Promise<void>;
  selectedFormulaNode: () => FormulaNodeConfig | null;
}

function extractErrorMessage(error: unknown, fallback: string): string {
  if (
    error &&
    typeof error === "object" &&
    "response" in error &&
    (error as any).response?.data
  ) {
    const payload = (error as any).response.data;
    if (typeof payload === "string") return payload;
    if (payload?.message) return String(payload.message);
    if (payload?.detail) return String(payload.detail);
  }
  return error instanceof Error ? error.message : fallback;
}

export const useWorkbookStore = create<WorkbookState>((set, get) => ({
  workbook: null,
  versions: [],
  selectedNodeId: null,
  loading: false,
  error: null,

  async uploadFiles(payload) {
    set({ loading: true, error: null });
    try {
      const response = await uploadWorkbook({
        workbookId: get().workbook?.workbookId,
        inputFile: payload.inputFile,
        outputFile: payload.outputFile,
      });

      set({
        workbook: response.workbook,
        versions: response.versions,
        selectedNodeId: null,
        loading: false,
      });
    } catch (error) {
      set({
        loading: false,
        error: extractErrorMessage(error, "Upload failed"),
      });
    }
  },

  setSelectedNode(id) {
    set({ selectedNodeId: id });
  },

  async updateFormulaNode(update, label) {
    const workbookId = get().workbook?.workbookId;
    if (!workbookId) {
      return false;
    }

    set({ loading: true, error: null });
    try {
      const response = await recomputeWorkbook(
        workbookId,
        [update],
        label ?? "Edit formula node",
      );
      set({
        workbook: response.workbook,
        versions: response.versions,
        loading: false,
      });
      return true;
    } catch (error) {
      set({
        loading: false,
        error: extractErrorMessage(error, "Recompute failed"),
      });
      return false;
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
        loading: false,
      });
    } catch (error) {
      set({ loading: false, error: extractErrorMessage(error, "Undo failed") });
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
        loading: false,
      });
    } catch (error) {
      set({ loading: false, error: extractErrorMessage(error, "Redo failed") });
    }
  },

  async runPipeline(label?: string) {
    const workbookId = get().workbook?.workbookId;
    if (!workbookId) {
      return;
    }

    set({ loading: true, error: null });
    try {
      const response = await runPipelineRequest(workbookId, label);
      set({
        workbook: response.workbook,
        versions: response.versions,
        loading: false,
      });
    } catch (error) {
      set({
        loading: false,
        error: error instanceof Error ? error.message : "Pipeline run failed",
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
      anchor.download = `pipeline-${workbook.workbookId}.xlsx`;
      anchor.click();
      URL.revokeObjectURL(url);
      set({ loading: false });
    } catch (error) {
      set({
        loading: false,
        error: extractErrorMessage(error, "Export failed"),
      });
    }
  },

  selectedFormulaNode() {
    const selected = get().selectedNodeId;
    if (!selected) {
      return null;
    }
    return (
      get().workbook?.config.formulas.find((item) => item.id === selected) ??
      null
    );
  },
}));

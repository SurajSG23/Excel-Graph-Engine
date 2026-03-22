import { create } from "zustand";
import { exportWorkbook, recomputeWorkbook, uploadWorkbook } from "../services/api";
import { NodeUpdate, VersionItem, WorkbookGraph } from "../types/workbook";

interface WorkbookState {
  workbook: WorkbookGraph | null;
  versions: VersionItem[];
  selectedNodeId: string | null;
  selectedSheet: string | "ALL";
  searchText: string;
  showZeroDependencyNodes: boolean;
  loading: boolean;
  error: string | null;
  uploadFile: (file: File) => Promise<void>;
  setSelectedNode: (id: string | null) => void;
  setSelectedSheet: (sheet: string | "ALL") => void;
  setSearchText: (value: string) => void;
  setShowZeroDependencyNodes: (value: boolean) => void;
  applyUpdate: (updates: NodeUpdate[], label?: string) => Promise<void>;
  triggerExport: () => Promise<void>;
}

export const useWorkbookStore = create<WorkbookState>((set, get) => ({
  workbook: null,
  versions: [],
  selectedNodeId: null,
  selectedSheet: "ALL",
  searchText: "",
  showZeroDependencyNodes: true,
  loading: false,
  error: null,

  async uploadFile(file) {
    set({ loading: true, error: null });
    try {
      const response = await uploadWorkbook(file);
      set({
        workbook: response.workbook,
        versions: response.versions,
        selectedNodeId: null,
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
  }
}));

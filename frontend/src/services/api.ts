import axios from "axios";
import { NodeUpdate, WorkbookOperation, WorkbookResponse } from "../types/workbook";

const api = axios.create({
  baseURL: "/api"
});

export interface UploadPayload {
  workbookId?: string;
  inputFile?: File;
  outputFile?: File;
  labeledFile?: File;
  role?: "input" | "output";
}

export async function uploadWorkbook(payload: UploadPayload): Promise<WorkbookResponse> {
  const form = new FormData();
  if (payload.workbookId) {
    form.append("workbookId", payload.workbookId);
  }

  if (payload.inputFile) {
    form.append("input", payload.inputFile);
  }

  if (payload.outputFile) {
    form.append("output", payload.outputFile);
  }

  if (payload.labeledFile) {
    form.append("file", payload.labeledFile);
  }

  if (payload.role) {
    form.append("role", payload.role);
  }

  const { data } = await api.post<WorkbookResponse>("/upload", form, {
    headers: {
      "Content-Type": "multipart/form-data"
    }
  });
  return data;
}

export async function recomputeWorkbook(workbookId: string, updates: NodeUpdate[], label?: string): Promise<WorkbookResponse> {
  const { data } = await api.post<WorkbookResponse>("/recompute", {
    workbookId,
    updates,
    label
  });
  return data;
}

export async function runPipeline(workbookId: string, label?: string): Promise<WorkbookResponse> {
  const { data } = await api.post<WorkbookResponse>("/run", {
    workbookId,
    label
  });
  return data;
}

export async function exportWorkbook(workbookId: string): Promise<Blob> {
  const { data } = await api.post("/export", { workbookId }, { responseType: "blob" });
  return data as Blob;
}

export async function applyWorkbookOperations(
  workbookId: string,
  operations: WorkbookOperation[],
  label?: string
): Promise<WorkbookResponse> {
  const { data } = await api.post<WorkbookResponse>("/operations", {
    workbookId,
    operations,
    label
  });
  return data;
}

export async function undoWorkbook(workbookId: string): Promise<WorkbookResponse> {
  const { data } = await api.post<WorkbookResponse>("/undo", { workbookId });
  return data;
}

export async function redoWorkbook(workbookId: string): Promise<WorkbookResponse> {
  const { data } = await api.post<WorkbookResponse>("/redo", { workbookId });
  return data;
}

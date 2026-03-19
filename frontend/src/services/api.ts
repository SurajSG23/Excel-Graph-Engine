import axios from "axios";
import { NodeUpdate, WorkbookResponse } from "../types/workbook";

const api = axios.create({
  baseURL: "/api"
});

export async function uploadWorkbook(file: File): Promise<WorkbookResponse> {
  const form = new FormData();
  form.append("file", file);
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

export async function exportWorkbook(workbookId: string): Promise<Blob> {
  const { data } = await api.post("/export", { workbookId }, { responseType: "blob" });
  return data as Blob;
}

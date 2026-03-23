import { ChangeEvent, useState } from "react";
import { useWorkbookStore } from "../store/workbookStore";

export function UploadCard() {
  const [inputFile, setInputFile] = useState<File | undefined>();
  const [outputFile, setOutputFile] = useState<File | undefined>();
  const uploadFiles = useWorkbookStore((s) => s.uploadFiles);
  const loading = useWorkbookStore((s) => s.loading);

  const onInputChange = (event: ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0];
    setInputFile(file);
  };

  const onOutputChange = (event: ChangeEvent<HTMLInputElement>): void => {
    const file = event.target.files?.[0];
    setOutputFile(file);
  };

  const onUploadPair = async (): Promise<void> => {
    if (!inputFile && !outputFile) {
      return;
    }

    await uploadFiles({ inputFile, outputFile });
  };

  const onUploadLabeled = async (role: "input" | "output"): Promise<void> => {
    const file = role === "input" ? inputFile : outputFile;
    if (!file) {
      return;
    }

    await uploadFiles({ labeledFile: file, role });
  };

  return (
    <section className="upload-card panel">
      <summary>Workbook Upload</summary>
      <p>Provide input and output workbooks together, or upload one side at a time.</p>

      <div className="upload-grid">
        <label className={`upload-dropzone ${loading ? "is-loading" : ""}`}>
          <input
            className="upload-input-native"
            type="file"
            accept=".xlsx"
            onChange={onInputChange}
            disabled={loading}
          />
          <span className="upload-dropzone-line upload-dropzone-title">Input workbook</span>
          <span className="upload-dropzone-line upload-dropzone-subtitle">
            {inputFile ? inputFile.name : "Select .xlsx"}
          </span>
        </label>

        <label className={`upload-dropzone ${loading ? "is-loading" : ""}`}>
          <input
            className="upload-input-native"
            type="file"
            accept=".xlsx"
            onChange={onOutputChange}
            disabled={loading}
          />
          <span className="upload-dropzone-line upload-dropzone-title">Output workbook</span>
          <span className="upload-dropzone-line upload-dropzone-subtitle">
            {outputFile ? outputFile.name : "Select .xlsx"}
          </span>
        </label>
      </div>

      <div className="upload-actions">
        <button type="button" onClick={onUploadPair} disabled={loading || (!inputFile && !outputFile)}>
          {loading ? "Processing..." : "Upload Selected Workbooks"}
        </button>
        <button type="button" onClick={() => onUploadLabeled("input")} disabled={loading || !inputFile}>
          Upload Input Only
        </button>
        <button type="button" onClick={() => onUploadLabeled("output")} disabled={loading || !outputFile}>
          Upload Output Only
        </button>
      </div>
    </section>
  );
}

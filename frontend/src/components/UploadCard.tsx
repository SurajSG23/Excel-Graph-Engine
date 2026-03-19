import { ChangeEvent, useState } from "react";
import { useWorkbookStore } from "../store/workbookStore";

export function UploadCard() {
  const [selectedFileName, setSelectedFileName] = useState<string>("");
  const uploadFile = useWorkbookStore((s) => s.uploadFile);
  const loading = useWorkbookStore((s) => s.loading);

  const onFileChange = async (event: ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = event.target.files?.[0];
    if (!file) return;
    setSelectedFileName(file.name);
    await uploadFile(file);
  };

  return (
    <section className="upload-card">
      <h1>Excel Graph Engine</h1>
      <p>Upload an .xlsx workbook and convert every sheet into one connected dependency graph.</p>
      <label className="upload-input">
        <input type="file" accept=".xlsx" onChange={onFileChange} disabled={loading} />
        <span>{loading ? "Processing workbook..." : "Select workbook"}</span>
      </label>
      {selectedFileName && <small>Loaded: {selectedFileName}</small>}
    </section>
  );
}

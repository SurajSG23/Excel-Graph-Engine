import { useWorkbookStore } from "../store/workbookStore";
import { useMemo } from "react";
import { Redo2, Undo2 } from "lucide-react";

export function TopBar() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedFile = useWorkbookStore((s) => s.selectedFile);
  const selectedSheet = useWorkbookStore((s) => s.selectedSheet);
  const searchText = useWorkbookStore((s) => s.searchText);
  const showZeroDependencyNodes = useWorkbookStore(
    (s) => s.showZeroDependencyNodes,
  );
  const setSelectedSheet = useWorkbookStore((s) => s.setSelectedSheet);
  const setSelectedFile = useWorkbookStore((s) => s.setSelectedFile);
  const setSearchText = useWorkbookStore((s) => s.setSearchText);
  const setShowZeroDependencyNodes = useWorkbookStore(
    (s) => s.setShowZeroDependencyNodes,
  );
  const undo = useWorkbookStore((s) => s.undo);
  const redo = useWorkbookStore((s) => s.redo);
  const triggerExport = useWorkbookStore((s) => s.triggerExport);
  const sheets = workbook
    ? selectedFile === "ALL"
      ? workbook.files.flatMap((file) => file.sheets.map((sheet) => `${file.fileName}::${sheet}`))
      : workbook.files
          .filter((file) => file.fileName === selectedFile)
          .flatMap((file) => file.sheets.map((sheet) => `${file.fileName}::${sheet}`))
    : [];

  return (
    <section className="panel topbar">
      <details className="panel-collapsible" open>
        <summary>Graph Controls</summary>
        <div className="toolbar-group">
          <label>
            File Selector
            <select value={selectedFile} onChange={(e) => setSelectedFile(e.target.value)}>
              <option value="ALL">All Files</option>
              {workbook?.files.map((file) => (
                <option key={file.fileName} value={file.fileName}>
                  {file.fileName} ({file.role})
                </option>
              ))}
            </select>
          </label>

          <label>
            Sheet Selector
            <select value={selectedSheet} onChange={(e) => setSelectedSheet(e.target.value)}>
              <option value="ALL">All Sheets</option>
              {sheets.map((sheet) => (
                <option key={sheet} value={sheet}>
                  {sheet}
                </option>
              ))}
            </select>
          </label>

          <label>
            Search Nodes
            <input
              value={searchText}
              placeholder="Node, cell, formula"
              onChange={(e) => setSearchText(e.target.value)}
            />
          </label>

          <label className="toggle-row">
            <input
              className="toggle-input"
              type="checkbox"
              checked={showZeroDependencyNodes}
              onChange={(e) => setShowZeroDependencyNodes(e.target.checked)}
            />
            <span className="toggle-text">Show isolated nodes</span>
          </label>
        </div>

        

        

        <div className="toolbar-group toolbar-inline-actions">
          <button
            type="button"
            className="toolbar-icon-button"
            onClick={() => undo()}
            disabled={!workbook}
            aria-label="Undo"
            title="Undo"
          >
            <Undo2 size={16} strokeWidth={2.2} aria-hidden="true" />
            <span>Undo</span>
          </button>
          <button
            type="button"
            className="toolbar-icon-button"
            onClick={() => redo()}
            disabled={!workbook}
            aria-label="Redo"
            title="Redo"
          >
            <Redo2 size={16} strokeWidth={2.2} aria-hidden="true" />
            <span>Redo</span>
          </button>
        </div>

        <button onClick={() => triggerExport()} disabled={!workbook}>
          Export Workbook
        </button>
      </details>
    </section>
  );
}

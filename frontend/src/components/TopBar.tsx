import { useWorkbookStore } from "../store/workbookStore";
import { useMemo, useState } from "react";

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
  const applyOperations = useWorkbookStore((s) => s.applyOperations);
  const undo = useWorkbookStore((s) => s.undo);
  const redo = useWorkbookStore((s) => s.redo);
  const triggerExport = useWorkbookStore((s) => s.triggerExport);
  const [rowIndex, setRowIndex] = useState("1");
  const [colIndex, setColIndex] = useState("1");
  const [newSheetName, setNewSheetName] = useState("NewSheet");
  const [renameSheetName, setRenameSheetName] = useState("RenamedSheet");

  const activeFile = selectedFile === "ALL" ? workbook?.files[0]?.fileName : selectedFile;
  const activeSheet = useMemo(() => {
    if (!workbook) {
      return undefined;
    }

    if (selectedSheet !== "ALL") {
      return selectedSheet.split("::")[1];
    }

    if (!activeFile) {
      return undefined;
    }

    const file = workbook.files.find((entry) => entry.fileName === activeFile);
    return file?.sheets[0];
  }, [workbook, selectedSheet, activeFile]);

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

        <div className="toolbar-group">
          <label>
            Row Index
            <input value={rowIndex} onChange={(e) => setRowIndex(e.target.value)} placeholder="1" />
          </label>
          <label>
            Column Index
            <input value={colIndex} onChange={(e) => setColIndex(e.target.value)} placeholder="1" />
          </label>
          <button
            type="button"
            disabled={!activeFile || !activeSheet}
            onClick={() =>
              applyOperations(
                [{ type: "INSERT_ROW", fileName: activeFile!, sheet: activeSheet!, index: Number(rowIndex) || 1 }],
                "Insert row"
              )
            }
          >
            Insert Row
          </button>
          <button
            type="button"
            disabled={!activeFile || !activeSheet}
            onClick={() =>
              applyOperations(
                [{ type: "DELETE_ROW", fileName: activeFile!, sheet: activeSheet!, index: Number(rowIndex) || 1 }],
                "Delete row"
              )
            }
          >
            Delete Row
          </button>
          <button
            type="button"
            disabled={!activeFile || !activeSheet}
            onClick={() =>
              applyOperations(
                [{ type: "INSERT_COLUMN", fileName: activeFile!, sheet: activeSheet!, index: Number(colIndex) || 1 }],
                "Insert column"
              )
            }
          >
            Insert Column
          </button>
          <button
            type="button"
            disabled={!activeFile || !activeSheet}
            onClick={() =>
              applyOperations(
                [{ type: "DELETE_COLUMN", fileName: activeFile!, sheet: activeSheet!, index: Number(colIndex) || 1 }],
                "Delete column"
              )
            }
          >
            Delete Column
          </button>
        </div>

        <div className="toolbar-group">
          <label>
            New Sheet Name
            <input value={newSheetName} onChange={(e) => setNewSheetName(e.target.value)} />
          </label>
          <button
            type="button"
            disabled={!activeFile || !newSheetName.trim()}
            onClick={() =>
              applyOperations(
                [{ type: "ADD_SHEET", fileName: activeFile!, sheet: newSheetName.trim() }],
                "Add sheet"
              )
            }
          >
            Add Sheet
          </button>
          <label>
            Rename Active Sheet To
            <input value={renameSheetName} onChange={(e) => setRenameSheetName(e.target.value)} />
          </label>
          <button
            type="button"
            disabled={!activeFile || !activeSheet || !renameSheetName.trim()}
            onClick={() =>
              applyOperations(
                [
                  {
                    type: "RENAME_SHEET",
                    fileName: activeFile!,
                    fromSheet: activeSheet!,
                    toSheet: renameSheetName.trim()
                  }
                ],
                "Rename sheet"
              )
            }
          >
            Rename Sheet
          </button>
          <button
            type="button"
            disabled={!activeFile || !activeSheet}
            onClick={() =>
              applyOperations(
                [{ type: "DELETE_SHEET", fileName: activeFile!, sheet: activeSheet! }],
                "Delete sheet"
              )
            }
          >
            Delete Sheet
          </button>
        </div>

        <div className="toolbar-group">
          <button type="button" onClick={() => undo()} disabled={!workbook}>
            Undo
          </button>
          <button type="button" onClick={() => redo()} disabled={!workbook}>
            Redo
          </button>
        </div>

        <button onClick={() => triggerExport()} disabled={!workbook}>
          Export Workbook
        </button>
      </details>
    </section>
  );
}

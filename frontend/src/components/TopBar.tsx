import { useWorkbookStore } from "../store/workbookStore";

export function TopBar() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedSheet = useWorkbookStore((s) => s.selectedSheet);
  const searchText = useWorkbookStore((s) => s.searchText);
  const setSelectedSheet = useWorkbookStore((s) => s.setSelectedSheet);
  const setSearchText = useWorkbookStore((s) => s.setSearchText);
  const triggerExport = useWorkbookStore((s) => s.triggerExport);

  return (
    <header className="topbar">
      <div className="toolbar-group">
        <label>
          Sheet
          <select value={selectedSheet} onChange={(e) => setSelectedSheet(e.target.value)}>
            <option value="ALL">All Sheets</option>
            {workbook?.sheets.map((sheet) => (
              <option key={sheet} value={sheet}>
                {sheet}
              </option>
            ))}
          </select>
        </label>

        <label>
          Search
          <input
            value={searchText}
            placeholder="Node, cell, formula"
            onChange={(e) => setSearchText(e.target.value)}
          />
        </label>
      </div>

      <button onClick={() => triggerExport()} disabled={!workbook}>
        Export Workbook
      </button>
    </header>
  );
}

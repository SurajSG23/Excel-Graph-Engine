import { useWorkbookStore } from "../store/workbookStore";

export function TopBar() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedSheet = useWorkbookStore((s) => s.selectedSheet);
  const searchText = useWorkbookStore((s) => s.searchText);
  const showZeroDependencyNodes = useWorkbookStore((s) => s.showZeroDependencyNodes);
  const setSelectedSheet = useWorkbookStore((s) => s.setSelectedSheet);
  const setSearchText = useWorkbookStore((s) => s.setSearchText);
  const setShowZeroDependencyNodes = useWorkbookStore((s) => s.setShowZeroDependencyNodes);
  const triggerExport = useWorkbookStore((s) => s.triggerExport);

  return (
    <section className="panel topbar">
      <details className="panel-collapsible" open>
        <summary>Graph Controls</summary>
        <div className="toolbar-group">
          <label>
            Sheet Selector
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
            Search Nodes
            <input
              value={searchText}
              placeholder="Node, cell, formula"
              onChange={(e) => setSearchText(e.target.value)}
            />
          </label>

          <label className="toggle-row">
            <input
              type="checkbox"
              checked={showZeroDependencyNodes}
              onChange={(e) => setShowZeroDependencyNodes(e.target.checked)}
            />
            Show 0-dependency nodes
          </label>
        </div>

        <button onClick={() => triggerExport()} disabled={!workbook}>
          Export Workbook
        </button>
      </details>
    </section>
  );
}

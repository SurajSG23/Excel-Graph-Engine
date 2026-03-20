import { UploadCard } from "./UploadCard";
import { TopBar } from "./TopBar";
import { NodeInspector } from "./NodeInspector";
import { ValidationPanel } from "./ValidationPanel";
import { VersionPanel } from "./VersionPanel";
import { useWorkbookStore } from "../store/workbookStore";

interface SidebarProps {
  isOpen: boolean;
  onToggle: () => void;
}

export function Sidebar({ isOpen, onToggle }: SidebarProps) {
  const error = useWorkbookStore((s) => s.error);
  const loading = useWorkbookStore((s) => s.loading);

  return (
    <aside className={`dashboard-sidebar ${isOpen ? "is-open" : "is-closed"}`}>
      <div className="sidebar-header">
        <h1>Excel Graph Engine</h1>
        <button
          type="button"
          className="sidebar-toggle icon-button"
          onClick={onToggle}
          aria-label={isOpen ? "Collapse controls panel" : "Expand controls panel"}
        >
          <span className={isOpen ? "icon-chevron-left" : "icon-chevron-right"} aria-hidden="true" />
        </button>
      </div>

      <div className="sidebar-content">
        <UploadCard />
        <TopBar />
        {error && <div className="error-banner">{error}</div>}
        {loading && <div className="loading-banner">Running workbook engine...</div>}
        <NodeInspector />
        <ValidationPanel />
        <VersionPanel />
      </div>
    </aside>
  );
}

import { useState } from "react";
import { GraphCanvas } from "./components/GraphCanvas";
import { Sidebar } from "./components/Sidebar";

function App() {
  const [sidebarOpen, setSidebarOpen] = useState(true);

  return (
    <div className={`app-shell dashboard-shell ${sidebarOpen ? "sidebar-open" : "sidebar-closed"}`}>
      <div className="grain" />
      <Sidebar isOpen={sidebarOpen} onToggle={() => setSidebarOpen((current) => !current)} />
      <button
        type="button"
        className="sidebar-backdrop"
        aria-label="Close controls panel"
        onClick={() => setSidebarOpen(false)}
      />
      <main className="dashboard-main">
        <button
          type="button"
          className="sidebar-fab icon-button"
          onClick={() => setSidebarOpen((current) => !current)}
          aria-label={sidebarOpen ? "Hide controls panel" : "Show controls panel"}
        >
          <span className={sidebarOpen ? "icon-chevron-left" : "icon-menu"} aria-hidden="true" />
        </button>
        <GraphCanvas />
      </main>
    </div>
  );
}

export default App;

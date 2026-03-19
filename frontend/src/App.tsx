import { UploadCard } from "./components/UploadCard";
import { TopBar } from "./components/TopBar";
import { GraphCanvas } from "./components/GraphCanvas";
import { NodeInspector } from "./components/NodeInspector";
import { ValidationPanel } from "./components/ValidationPanel";
import { VersionPanel } from "./components/VersionPanel";
import { useWorkbookStore } from "./store/workbookStore";

function App() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const loading = useWorkbookStore((s) => s.loading);
  const error = useWorkbookStore((s) => s.error);

  return (
    <div className="app-shell">
      <div className="grain" />
      <main>
        <UploadCard />
        {workbook && <TopBar />}

        {error && <div className="error-banner">{error}</div>}
        {loading && <div className="loading-banner">Running workbook engine...</div>}

        <section className="workspace-grid">
          <GraphCanvas />
          <aside className="side-panels">
            <NodeInspector />
            <ValidationPanel />
            <VersionPanel />
          </aside>
        </section>
      </main>
    </div>
  );
}

export default App;

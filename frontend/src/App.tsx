import { useEffect, useState } from "react";
import { GraphCanvas } from "./components/GraphCanvas";
import { Sidebar } from "./components/Sidebar";

function App() {
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [sidebarWidth, setSidebarWidth] = useState(380);
  const [isResizingSidebar, setIsResizingSidebar] = useState(false);

  useEffect(() => {
    if (!isResizingSidebar) {
      return;
    }

    const minWidth = 300;
    const maxWidth = 680;
    const onPointerMove = (event: PointerEvent) => {
      const nextWidth = Math.min(maxWidth, Math.max(minWidth, event.clientX));
      setSidebarWidth(nextWidth);
    };

    const onPointerUp = () => {
      setIsResizingSidebar(false);
    };

    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
    return () => {
      window.removeEventListener("pointermove", onPointerMove);
      window.removeEventListener("pointerup", onPointerUp);
    };
  }, [isResizingSidebar]);

  return (
    <div
      className={`app-shell dashboard-shell ${sidebarOpen ? "sidebar-open" : "sidebar-closed"} ${isResizingSidebar ? "sidebar-resizing" : ""}`}
      style={{ ["--sidebar-width" as string]: `${sidebarWidth}px` }}
    >
      <div className="grain" />
      <Sidebar
        isOpen={sidebarOpen}
        onToggle={() => setSidebarOpen((current) => !current)}
        onResizeStart={() => setIsResizingSidebar(true)}
      />
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

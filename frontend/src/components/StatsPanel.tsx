import { useMemo } from "react";
import { useWorkbookStore } from "../store/workbookStore";

export function StatsPanel() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedFile = useWorkbookStore((s) => s.selectedFile);
  const selectedSheet = useWorkbookStore((s) => s.selectedSheet);
  const searchText = useWorkbookStore((s) => s.searchText);

  const stats = useMemo(() => {
    if (!workbook) {
      return { nodes: 0, edges: 0, errors: 0, cycles: 0 };
    }

    const fileNodes =
      selectedFile === "ALL"
        ? workbook.nodes
        : workbook.nodes.filter((node) => node.fileName === selectedFile);

    const sheetNodes =
      selectedSheet === "ALL"
        ? fileNodes
        : fileNodes.filter((node) => `${node.fileName}::${node.sheet}` === selectedSheet);

    const query = searchText.trim().toLowerCase();
    const visibleNodes = query
      ? sheetNodes.filter(
          (node) =>
            node.id.toLowerCase().includes(query) ||
            node.cell.toLowerCase().includes(query) ||
            (node.formula ?? "").toLowerCase().includes(query),
        )
      : sheetNodes;

    const idSet = new Set(visibleNodes.map((node) => node.id));
    const visibleEdges = workbook.edges.filter(
      (edge) => idSet.has(edge.source) && idSet.has(edge.target),
    );

    const issues = workbook.validationIssues ?? [];
    const cycleCount = issues.filter(
      (issue) => issue.type === "CIRCULAR_DEPENDENCY",
    ).length;

    return {
      nodes: visibleNodes.length,
      edges: visibleEdges.length,
      errors: issues.length,
      cycles: cycleCount,
    };
  }, [workbook, selectedFile, selectedSheet, searchText]);

  return (
    <section className="panel stats-panel">
      <h3>Stats</h3>
      <div className="stats-grid">
        <div>
          <span>Nodes</span>
          <strong>{stats.nodes}</strong>
        </div>
        <div>
          <span>Edges</span>
          <strong>{stats.edges}</strong>
        </div>
        <div>
          <span>Errors</span>
          <strong>{stats.errors}</strong>
        </div>
        <div>
          <span>Cycles</span>
          <strong>{stats.cycles}</strong>
        </div>
      </div>
    </section>
  );
}

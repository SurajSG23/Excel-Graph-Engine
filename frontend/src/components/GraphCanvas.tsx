import { useMemo } from "react";
import { Background, Controls, MiniMap, Panel, ReactFlow } from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import { useWorkbookStore } from "../store/workbookStore";
import { buildTraversalSets, toFlowEdges, toFlowNodes } from "../utils/graphLayout";
import { CellNode } from "./CellNode";

const nodeTypes = {
  cellNode: CellNode
};

export function GraphCanvas() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedNodeId = useWorkbookStore((s) => s.selectedNodeId);
  const selectedSheet = useWorkbookStore((s) => s.selectedSheet);
  const searchText = useWorkbookStore((s) => s.searchText);
  const setSelectedNode = useWorkbookStore((s) => s.setSelectedNode);

  const filtered = useMemo(() => {
    if (!workbook) return { nodes: [], edges: [] };

    const sheetNodes = selectedSheet === "ALL"
      ? workbook.nodes
      : workbook.nodes.filter((node) => node.sheet === selectedSheet);

    const query = searchText.trim().toLowerCase();
    const visibleNodes = query
      ? sheetNodes.filter(
          (node) =>
            node.id.toLowerCase().includes(query) ||
            node.cell.toLowerCase().includes(query) ||
            (node.formula ?? "").toLowerCase().includes(query)
        )
      : sheetNodes;

    const idSet = new Set(visibleNodes.map((n) => n.id));
    const visibleEdges = workbook.edges.filter((edge) => idSet.has(edge.source) && idSet.has(edge.target));

    return { nodes: visibleNodes, edges: visibleEdges };
  }, [searchText, selectedSheet, workbook]);

  const highlight = useMemo(() => {
    if (!selectedNodeId || !workbook) return new Set<string>();
    const { upstream, downstream } = buildTraversalSets(selectedNodeId, workbook.edges);
    return new Set<string>([selectedNodeId, ...upstream, ...downstream]);
  }, [selectedNodeId, workbook]);

  const flowNodes = useMemo(
    () => toFlowNodes(filtered.nodes, highlight, selectedNodeId),
    [filtered.nodes, highlight, selectedNodeId]
  );

  const flowEdges = useMemo(
    () => toFlowEdges(filtered.edges, highlight),
    [filtered.edges, highlight]
  );

  if (!workbook) {
    return (
      <section className="canvas-empty">
        <h2>No workbook loaded</h2>
        <p>Upload a workbook to start visualizing cell dependencies.</p>
      </section>
    );
  }

  return (
    <section className="canvas-shell">
      <ReactFlow
        nodes={flowNodes}
        edges={flowEdges}
        nodeTypes={nodeTypes}
        minZoom={0.25}
        maxZoom={1.8}
        fitViewOptions={{ padding: 0.2 }}
        fitView
        onNodeClick={(_, node) => setSelectedNode(node.id)}
      >
        <MiniMap
          pannable
          zoomable
          nodeBorderRadius={8}
          nodeColor={(node) => String((node.data as { color?: string } | undefined)?.color ?? "#94a3b8")}
          maskColor="rgba(22,101,52,0.08)"
        />
        <Controls />
        <Panel position="top-left">
          <div className="graph-help-card">
            <h4>Workbook Graph</h4>
            <p>Select a node to trace upstream and downstream impact.</p>
            <div className="graph-help-stats">
              <span>{filtered.nodes.length} nodes</span>
              <span>{filtered.edges.length} edges</span>
            </div>
          </div>
        </Panel>
        <Background gap={28} size={1.1} color="#d5e7dc" />
      </ReactFlow>
    </section>
  );
}

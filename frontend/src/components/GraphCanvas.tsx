import { useEffect, useMemo, useRef, useState } from "react";
import {
  Background,
  Controls,
  MiniMap,
  Panel,
  ReactFlow,
  ReactFlowInstance,
} from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import { useWorkbookStore } from "../store/workbookStore";
import {
  buildTraversalSets,
  toFlowEdges,
  toFlowNodes,
} from "../utils/graphLayout";
import { CellNode } from "./CellNode";
import { SheetGroupNode } from "./SheetGroupNode";

const nodeTypes = {
  cellNode: CellNode,
  sheetGroup: SheetGroupNode,
};

export function GraphCanvas() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedNodeId = useWorkbookStore((s) => s.selectedNodeId);
  const selectedSheet = useWorkbookStore((s) => s.selectedSheet);
  const searchText = useWorkbookStore((s) => s.searchText);
  const showZeroDependencyNodes = useWorkbookStore((s) => s.showZeroDependencyNodes);
  const setSelectedNode = useWorkbookStore((s) => s.setSelectedNode);
  const [hoveredNodeId, setHoveredNodeId] = useState<string | null>(null);
  const [zoomLevel, setZoomLevel] = useState(1);
  const [reactFlowInstance, setReactFlowInstance] =
    useState<ReactFlowInstance | null>(null);
  const lastFitKey = useRef<string>("");

  const filtered = useMemo(() => {
    if (!workbook) return { nodes: [], edges: [] };

    const sheetNodes =
      selectedSheet === "ALL"
        ? workbook.nodes
        : workbook.nodes.filter((node) => node.sheet === selectedSheet);

    const query = searchText.trim().toLowerCase();
    const queryFilteredNodes = query
      ? sheetNodes.filter(
          (node) =>
            node.id.toLowerCase().includes(query) ||
            node.cell.toLowerCase().includes(query) ||
            (node.formula ?? "").toLowerCase().includes(query),
        )
      : sheetNodes;

    const candidateIdSet = new Set(queryFilteredNodes.map((node) => node.id));
    const candidateEdges = workbook.edges.filter(
      (edge) => candidateIdSet.has(edge.source) && candidateIdSet.has(edge.target),
    );

    const connectedNodeIds = new Set<string>();
    for (const edge of candidateEdges) {
      connectedNodeIds.add(edge.source);
      connectedNodeIds.add(edge.target);
    }

    const visibleNodes = showZeroDependencyNodes
      ? queryFilteredNodes
      : queryFilteredNodes.filter((node) => connectedNodeIds.has(node.id));

    const idSet = new Set(visibleNodes.map((n) => n.id));
    const visibleEdges = candidateEdges.filter(
      (edge) => idSet.has(edge.source) && idSet.has(edge.target),
    );

    return { nodes: visibleNodes, edges: visibleEdges };
  }, [searchText, selectedSheet, showZeroDependencyNodes, workbook]);

  const activeNodeId = selectedNodeId;

  const highlight = useMemo(() => {
    if (!activeNodeId || !workbook) return new Set<string>();
    const selectedNode = workbook.nodes.find((node) => node.id === activeNodeId);
    return new Set<string>([activeNodeId, ...(selectedNode?.dependencies ?? [])]);
  }, [activeNodeId, workbook]);

  const traversal = useMemo(() => {
    if (!activeNodeId || !workbook) {
      return { upstream: new Set<string>(), downstream: new Set<string>() };
    }
    const selectedNode = workbook.nodes.find((node) => node.id === activeNodeId);
    return {
      upstream: new Set(selectedNode?.dependencies ?? []),
      downstream: new Set<string>()
    };
  }, [activeNodeId, workbook]);

  const issueSummary = useMemo(() => {
    const allIssues = workbook?.validationIssues ?? [];
    const errorNodeIds = new Set<string>();
    const circularNodeIds = new Set<string>();

    for (const issue of allIssues) {
      if (issue.nodeId) {
        errorNodeIds.add(issue.nodeId);
      }
      for (const related of issue.relatedNodeIds ?? []) {
        errorNodeIds.add(related);
        if (issue.type === "CIRCULAR_DEPENDENCY") {
          circularNodeIds.add(related);
        }
      }
      if (issue.type === "CIRCULAR_DEPENDENCY" && issue.nodeId) {
        circularNodeIds.add(issue.nodeId);
      }
    }

    return {
      errorNodeIds,
      circularNodeIds,
      errorCount: allIssues.length,
      circularCount: allIssues.filter((i) => i.type === "CIRCULAR_DEPENDENCY")
        .length,
    };
  }, [workbook]);

  const nodeSheetMap = useMemo(
    () => new Map((workbook?.nodes ?? []).map((node) => [node.id, node.sheet])),
    [workbook],
  );

  const flowNodes = useMemo(
    () =>
      toFlowNodes(filtered.nodes, filtered.edges, {
        selectedNodeId,
        highlight,
        upstream: traversal.upstream,
        downstream: traversal.downstream,
        errorNodeIds: issueSummary.errorNodeIds,
        circularNodeIds: issueSummary.circularNodeIds,
        selectedSheet,
        zoomLevel,
      }),
    [
      filtered.nodes,
      filtered.edges,
      selectedNodeId,
      highlight,
      traversal.upstream,
      traversal.downstream,
      issueSummary.errorNodeIds,
      issueSummary.circularNodeIds,
      selectedSheet,
      zoomLevel,
    ],
  );

  const flowEdges = useMemo(
    () =>
      toFlowEdges(
        filtered.edges,
        highlight,
        nodeSheetMap,
        selectedNodeId,
        hoveredNodeId,
      ),
    [filtered.edges, highlight, nodeSheetMap, selectedNodeId, hoveredNodeId],
  );

  const [nodes, setNodes] = useState(flowNodes);
  const [edges, setEdges] = useState(flowEdges);

  useEffect(() => {
    setNodes(flowNodes);
  }, [flowNodes, setNodes]);

  useEffect(() => {
    setEdges(flowEdges);
  }, [flowEdges, setEdges]);

  useEffect(() => {
    if (!workbook) {
      return;
    }

    const fitKey = `${selectedSheet}|${searchText}|${showZeroDependencyNodes}|${filtered.nodes.length}|${filtered.edges.length}`;
    if (fitKey === lastFitKey.current) {
      return;
    }

    lastFitKey.current = fitKey;
    if (!reactFlowInstance) {
      return;
    }

    const handle = requestAnimationFrame(() => {
      reactFlowInstance.fitView({ padding: 0.2, duration: 260 });
    });

    return () => cancelAnimationFrame(handle);
  }, [
    reactFlowInstance,
    workbook,
    selectedSheet,
    searchText,
    showZeroDependencyNodes,
    filtered.nodes.length,
    filtered.edges.length,
  ]);

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
        nodes={nodes}
        edges={edges}
        nodeTypes={nodeTypes}
        fitView
        fitViewOptions={{ padding: 0.18 }}
        minZoom={0.25}
        maxZoom={1.8}
        onInit={setReactFlowInstance}
        onNodeClick={(_, node) => {
          if (node.type !== "cellNode") {
            return;
          }
          setSelectedNode(node.id);
        }}
        onNodeMouseEnter={(_, node) => {
          if (node.type === "cellNode") {
            setHoveredNodeId(node.id);
          }
        }}
        onNodeMouseLeave={() => setHoveredNodeId(null)}
        onPaneClick={() => {
          setSelectedNode(null);
          setHoveredNodeId(null);
        }}
        onMove={(_event, viewport) => setZoomLevel(viewport.zoom)}
      >
        <MiniMap
          pannable
          zoomable
          style={{ width: 130, height: 78, borderRadius: 8 }}
          nodeBorderRadius={8}
          nodeColor={(node) => {
            if (node.type === "sheetGroup") {
              return "#eaf4ee";
            }
            return String(
              (node.data as { roleColor?: string } | undefined)?.roleColor ??
                "#94a3b8",
            );
          }}
          maskColor="rgba(22,101,52,0.08)"
        />
        <Panel position="top-left">
          <div className="graph-help-card">
            <h4>Stats</h4>
            <div className="graph-help-stats">
              <div>
                <span>{nodes.length} nodes</span>
                <span>{edges.length} edges</span>
              </div>
              <div>
                <span>{issueSummary.errorCount} errors</span>
                <span>{issueSummary.circularCount} cycles</span>
              </div>
            </div>
          </div>
        </Panel>
        <Panel position="top-right">
          <div className="graph-legend-card">
            <h4>Legend</h4>
            <ul>
              <li>
                <i className="legend-dot input" /> Input
              </li>
              <li>
                <i className="legend-dot computed" /> Computed
              </li>
              <li>
                <i className="legend-dot output" /> Output
              </li>
              <li>
                <i className="legend-dot error" /> Error/Cycle
              </li>
              <li>
                <i className="legend-line same" /> Same-sheet
              </li>
              <li>
                <i className="legend-line cross" /> Cross-sheet
              </li>
            </ul>
          </div>
        </Panel>
        <Controls showInteractive={false} />
        <Background gap={28} size={1.1} color="#d5e7dc" />
      </ReactFlow>
    </section>
  );
}

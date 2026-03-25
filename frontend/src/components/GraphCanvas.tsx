import { useEffect, useMemo, useRef, useState } from "react";
import {
  Background,
  Controls,
  Edge,
  MiniMap,
  Node,
  Panel,
  ReactFlow,
  ReactFlowInstance,
  useEdgesState,
  useNodesState,
} from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import { useWorkbookStore } from "../store/workbookStore";
import {
  buildTraversalSets,
  FlowCellData,
  FlowFormulaGroupData,
  FlowRoleGroupData,
  FlowSheetGroupData,
  toFlowEdges,
  toFlowNodes,
} from "../utils/graphLayout";
import { CellNode } from "./CellNode";
import { SheetGroupNode } from "./SheetGroupNode";
import { RoleGroupNode } from "./RoleGroupNode";
import { FormulaGroupNode } from "./FormulaGroupNode";
import { isGroupedNode, projectGraphForFormulaGrouping } from "../utils/formulaGrouping";

const nodeTypes = {
  cellNode: CellNode,
  formulaGroup: FormulaGroupNode,
  roleGroup: RoleGroupNode,
  sheetGroup: SheetGroupNode,
};

type FlowNode =
  | Node<FlowCellData>
  | Node<FlowFormulaGroupData>
  | Node<FlowRoleGroupData>
  | Node<FlowSheetGroupData>;
type FlowEdge = Edge;

export function GraphCanvas() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedNodeId = useWorkbookStore((s) => s.selectedNodeId);
  const selectedFile = useWorkbookStore((s) => s.selectedFile);
  const selectedSheet = useWorkbookStore((s) => s.selectedSheet);
  const searchText = useWorkbookStore((s) => s.searchText);
  const showZeroDependencyNodes = useWorkbookStore(
    (s) => s.showZeroDependencyNodes,
  );
  const groupSimilarFormulas = useWorkbookStore((s) => s.groupSimilarFormulas);
  const setSelectedNode = useWorkbookStore((s) => s.setSelectedNode);
  const applyOperations = useWorkbookStore((s) => s.applyOperations);
  const [hoveredNodeId, setHoveredNodeId] = useState<string | null>(null);
  const [reactFlowInstance, setReactFlowInstance] = useState<ReactFlowInstance<
    FlowNode,
    FlowEdge
  > | null>(null);
  const lastFitKey = useRef<string>("");

  const projectedGraph = useMemo(
    () => projectGraphForFormulaGrouping(workbook, groupSimilarFormulas),
    [workbook, groupSimilarFormulas],
  );

  const filtered = useMemo(() => {
    if (!workbook) return { nodes: [], edges: [] };

    const byFile =
      selectedFile === "ALL"
        ? projectedGraph.nodes
        : projectedGraph.nodes.filter((node) => node.fileName === selectedFile);

    const sheetNodes =
      selectedSheet === "ALL"
        ? byFile
        : byFile.filter(
            (node) => `${node.fileName}::${node.sheet}` === selectedSheet,
          );

    const query = searchText.trim().toLowerCase();
    const queryFilteredNodes = query
      ? sheetNodes.filter(
          (node) =>
            node.id.toLowerCase().includes(query) ||
            node.cell.toLowerCase().includes(query) ||
            (node.formula ?? "").toLowerCase().includes(query) ||
            (isGroupedNode(node) && node.formulaTemplate.toLowerCase().includes(query)),
        )
      : sheetNodes;

    const candidateIdSet = new Set(queryFilteredNodes.map((node) => node.id));
    const candidateEdges = projectedGraph.edges.filter(
      (edge) =>
        candidateIdSet.has(edge.source) && candidateIdSet.has(edge.target),
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
  }, [
    groupSimilarFormulas,
    projectedGraph.edges,
    projectedGraph.nodes,
    searchText,
    selectedFile,
    selectedSheet,
    showZeroDependencyNodes,
    workbook,
  ]);

  const activeNodeId = selectedNodeId;

  const highlight = useMemo(() => {
    if (!activeNodeId || !workbook) return new Set<string>();
    const selectedNode = projectedGraph.nodes.find(
      (node) => node.id === activeNodeId,
    );
    return new Set<string>([
      activeNodeId,
      ...(selectedNode?.dependencies ?? []),
    ]);
  }, [activeNodeId, projectedGraph.nodes, workbook]);

  const traversal = useMemo(() => {
    if (!activeNodeId || !workbook) {
      return { upstream: new Set<string>(), downstream: new Set<string>() };
    }
    return buildTraversalSets(activeNodeId, projectedGraph.edges);
  }, [activeNodeId, projectedGraph.edges, workbook]);

  const issueSummary = useMemo(() => {
    const allIssues = workbook?.validationIssues ?? [];
    const errorNodeIds = new Set<string>();
    const circularNodeIds = new Set<string>();

    for (const issue of allIssues) {
      if (issue.nodeId) {
        errorNodeIds.add(projectedGraph.nodeToGroupId.get(issue.nodeId) ?? issue.nodeId);
      }
      for (const related of issue.relatedNodeIds ?? []) {
        const mapped = projectedGraph.nodeToGroupId.get(related) ?? related;
        errorNodeIds.add(mapped);
        if (issue.type === "CIRCULAR_DEPENDENCY") {
          circularNodeIds.add(mapped);
        }
      }
      if (issue.type === "CIRCULAR_DEPENDENCY" && issue.nodeId) {
        circularNodeIds.add(projectedGraph.nodeToGroupId.get(issue.nodeId) ?? issue.nodeId);
      }
    }

    return {
      errorNodeIds,
      circularNodeIds,
      errorCount: allIssues.length,
      circularCount: allIssues.filter((i) => i.type === "CIRCULAR_DEPENDENCY")
        .length,
    };
  }, [projectedGraph.nodeToGroupId, workbook]);

  const nodeFileMap = useMemo(
    () =>
      new Map(projectedGraph.nodes.map((node) => [node.id, node.fileName])),
    [projectedGraph.nodes],
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
        selectedFile,
        selectedSheet,
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
      selectedFile,
      selectedSheet,
    ],
  );

  const sheetGroups = useMemo(
    () => flowNodes.filter((node) => node.type === "sheetGroup"),
    [flowNodes],
  );

  const flowEdges = useMemo(
    () =>
      toFlowEdges(
        filtered.edges,
        highlight,
        nodeFileMap,
        selectedNodeId,
        hoveredNodeId,
      ),
    [filtered.edges, highlight, nodeFileMap, selectedNodeId, hoveredNodeId],
  );

  const [nodes, setNodes, onNodesChange] = useNodesState<FlowNode>(
    flowNodes as FlowNode[],
  );
  const [edges, setEdges, onEdgesChange] = useEdgesState<FlowEdge>(
    flowEdges as FlowEdge[],
  );

  const stats = useMemo(
    () => ({
      nodeCount: filtered.nodes.length,
      edgeCount: filtered.edges.length,
    }),
    [filtered.nodes.length, filtered.edges.length],
  );

  useEffect(() => {
    setNodes(flowNodes);
  }, [flowNodes, setNodes]);

  useEffect(() => {
    setEdges(flowEdges);
  }, [flowEdges, setEdges]);

  useEffect(() => {
    if (!selectedNodeId) {
      return;
    }

    if (groupSimilarFormulas) {
      const mapped = projectedGraph.nodeToGroupId.get(selectedNodeId);
      if (mapped && mapped !== selectedNodeId) {
        setSelectedNode(mapped);
      }
      return;
    }

    if (selectedNodeId.startsWith("group:")) {
      setSelectedNode(null);
    }
  }, [groupSimilarFormulas, projectedGraph.nodeToGroupId, selectedNodeId, setSelectedNode]);

  useEffect(() => {
    if (!workbook) {
      return;
    }

    const fitKey = `${selectedFile}|${selectedSheet}|${searchText}|${showZeroDependencyNodes}|${groupSimilarFormulas}|${filtered.nodes.length}|${filtered.edges.length}`;
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
    selectedFile,
    selectedSheet,
    searchText,
    showZeroDependencyNodes,
    groupSimilarFormulas,
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
      <ReactFlow<FlowNode, FlowEdge>
        nodes={nodes}
        edges={edges}
        onNodesChange={onNodesChange}
        onEdgesChange={onEdgesChange}
        nodeTypes={nodeTypes}
        onlyRenderVisibleElements
        nodesDraggable
        fitView
        fitViewOptions={{ padding: 0.18 }}
        minZoom={0.25}
        maxZoom={1.8}
        onInit={setReactFlowInstance}
        onNodeClick={(_, node) => {
          if (node.type !== "cellNode" && node.type !== "formulaGroup") {
            return;
          }
          setSelectedNode(node.id);
        }}
        onNodeMouseEnter={(_, node) => {
          if (node.type === "cellNode" || node.type === "formulaGroup") {
            setHoveredNodeId(node.id);
          }
        }}
        onNodeMouseLeave={() => setHoveredNodeId(null)}
        onPaneClick={() => {
          setSelectedNode(null);
          setHoveredNodeId(null);
        }}
        onNodeDragStop={(_event, draggedNode) => {
          if (draggedNode.type !== "cellNode") {
            return;
          }

          const source = workbook.nodes.find(
            (node) => node.id === draggedNode.id,
          );
          if (!source) {
            return;
          }

          const width = Number(draggedNode.width ?? 150);
          const height = Number(draggedNode.height ?? 84);
          const centerX = draggedNode.position.x + width / 2;
          const centerY = draggedNode.position.y + height / 2;

          let targetGroup: Node | undefined;
          for (const group of sheetGroups) {
            const groupWidth = Number(group.style?.width ?? group.width ?? 0);
            const groupHeight = Number(
              group.style?.height ?? group.height ?? 0,
            );
            const minX = group.position.x;
            const minY = group.position.y;
            const maxX = minX + groupWidth;
            const maxY = minY + groupHeight;
            if (
              centerX >= minX &&
              centerX <= maxX &&
              centerY >= minY &&
              centerY <= maxY
            ) {
              targetGroup = group;
              break;
            }
          }

          if (!targetGroup) {
            return;
          }

          const targetData = targetGroup.data as
            | { fileName?: string; sheet?: string }
            | undefined;
          const targetFileName = targetData?.fileName;
          const targetSheet = targetData?.sheet;

          if (!targetFileName || !targetSheet) {
            return;
          }

          if (
            source.fileName === targetFileName &&
            source.sheet === targetSheet
          ) {
            return;
          }

          void applyOperations(
            [
              {
                type: "MOVE_CELL",
                fromNodeId: source.id,
                toFileName: targetFileName,
                toSheet: targetSheet,
                toCell: source.cell,
              },
            ],
            `Move ${source.id} to ${targetFileName}::${targetSheet}`,
          );
        }}
      >
        <MiniMap
          pannable
          zoomable
          style={{ width: 130, height: 78, borderRadius: 8 }}
          nodeBorderRadius={8}
          nodeColor={(node) => {
            if (node.type === "roleGroup") {
              return String(
                (node.data as { color?: string } | undefined)?.color ??
                  "#cbd5e1",
              );
            }
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
                <span>{stats.nodeCount} nodes</span>
                <span>{stats.edgeCount} edges</span>
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
              {groupSimilarFormulas && (
                <li>
                  <i className="legend-dot group" /> Grouped formula
                </li>
              )}
              <li>
                <i className="legend-line same" /> Same-file
              </li>
              <li>
                <i className="legend-line cross" /> Cross-file
              </li>
            </ul>
          </div>
        </Panel>
        <Controls showInteractive={false} />
        <Background id="dot-grid" gap={28} size={2} color="black" />
      </ReactFlow>
    </section>
  );
}

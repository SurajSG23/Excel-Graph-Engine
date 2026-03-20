import dagre from "dagre";
import { Edge, MarkerType, Node } from "@xyflow/react";
import { GraphEdge, GraphNode } from "../types/workbook";

const SHEET_COLORS = ["#16a34a", "#2563eb", "#9333ea", "#ea580c", "#0891b2", "#ca8a04", "#be123c"];
const ROLE_COLORS = {
  input: "#16a34a",
  computed: "#2563eb",
  output: "#9333ea",
  error: "#dc2626",
  circular: "#b91c1c"
} as const;

type Role = "input" | "computed" | "output" | "error" | "circular";

const NODE_WIDTH = 150;
const NODE_HEIGHT = 84;
const SHEETS_PER_ROW = 2;
const SHEET_PADDING_X = 42;
const SHEET_PADDING_Y = 56;
const SHEET_GAP_X = 90;
const SHEET_GAP_Y = 94;

export interface FlowCellData {
  [key: string]: unknown;
  label: string;
  id: string;
  sheet: string;
  value?: number;
  formula?: string;
  color: string;
  roleColor: string;
  role: Role;
  isSelected: boolean;
  isHighlighted: boolean;
  isUpstream: boolean;
  isDownstream: boolean;
  dependencyCount: number;
  dependencies: string[];
  isDimmed: boolean;
  isHovered: boolean;
  showExtra: boolean;
}

export interface FlowSheetGroupData {
  [key: string]: unknown;
  sheet: string;
  nodeCount: number;
  color: string;
}

interface FlowBuildContext {
  selectedNodeId: string | null;
  highlight: Set<string>;
  upstream: Set<string>;
  downstream: Set<string>;
  errorNodeIds: Set<string>;
  circularNodeIds: Set<string>;
  selectedSheet: string | "ALL";
  zoomLevel: number;
}

interface LayoutResult {
  positions: Map<string, { x: number; y: number }>;
  sheetBounds: Map<string, { x: number; y: number; width: number; height: number; nodeCount: number }>;
}

function hashSheetColor(sheet: string): string {
  let sum = 0;
  for (let i = 0; i < sheet.length; i += 1) {
    sum += sheet.charCodeAt(i);
  }
  return SHEET_COLORS[sum % SHEET_COLORS.length];
}

export function toFlowNodes(
  graphNodes: GraphNode[],
  graphEdges: GraphEdge[],
  context: FlowBuildContext
): Array<Node<FlowCellData> | Node<FlowSheetGroupData>> {
  const indegree = new Map<string, number>();
  const outdegree = new Map<string, number>();

  for (const node of graphNodes) {
    indegree.set(node.id, 0);
    outdegree.set(node.id, 0);
  }

  for (const edge of graphEdges) {
    indegree.set(edge.target, (indegree.get(edge.target) ?? 0) + 1);
    outdegree.set(edge.source, (outdegree.get(edge.source) ?? 0) + 1);
  }

  const layout = layoutBySheetDirectional(graphNodes, graphEdges, context.selectedSheet);

  const sheetGroups: Array<Node<FlowSheetGroupData>> = [...layout.sheetBounds.entries()].map(([sheet, bounds]) => ({
    id: `group:${sheet}`,
    type: "sheetGroup",
    position: { x: bounds.x, y: bounds.y },
    data: {
      sheet,
      nodeCount: bounds.nodeCount,
      color: hashSheetColor(sheet)
    },
    style: {
      width: bounds.width,
      height: bounds.height
    },
    selectable: false,
    draggable: false,
    connectable: false,
    zIndex: -1
  }));

  const cellNodes: Array<Node<FlowCellData>> = graphNodes.map((node) => {
    const color = hashSheetColor(node.sheet);
    const isSelected = context.selectedNodeId === node.id;
    const isHighlighted = context.highlight.has(node.id);
    const isCircular = context.circularNodeIds.has(node.id);
    const isError = context.errorNodeIds.has(node.id);
    const nodeIn = indegree.get(node.id) ?? 0;
    const nodeOut = outdegree.get(node.id) ?? 0;
    const hasSelection = Boolean(context.selectedNodeId);
    const isDimmed = hasSelection && !isHighlighted;

    let role: Role = node.formula ? "computed" : "input";
    if (nodeOut === 0 && nodeIn > 0) {
      role = "output";
    }
    if (isError) {
      role = isCircular ? "circular" : "error";
    }

    return {
      id: node.id,
      type: "cellNode",
      data: {
        id: node.id,
        label: `${node.cell}`,
        sheet: node.sheet,
        value: node.value,
        formula: node.formula,
        color,
        role,
        roleColor: ROLE_COLORS[role],
        isSelected,
        isHighlighted,
        isUpstream: context.upstream.has(node.id),
        isDownstream: context.downstream.has(node.id),
        dependencyCount: node.dependencies.length,
        dependencies: node.dependencies,
        isDimmed,
        isHovered: false,
        showExtra: context.zoomLevel > 0.95
      },
      position: layout.positions.get(node.id) ?? { x: 0, y: 0 },
      style: {
        width: role === "output" ? 170 : NODE_WIDTH,
        height: role === "output" ? 92 : NODE_HEIGHT
      }
    } satisfies Node<FlowCellData>;
  });

  return [...sheetGroups, ...cellNodes];
}

function layoutBySheetDirectional(
  graphNodes: GraphNode[],
  graphEdges: GraphEdge[],
  selectedSheet: string | "ALL"
): LayoutResult {
  const positions = new Map<string, { x: number; y: number }>();
  const sheetBounds = new Map<string, { x: number; y: number; width: number; height: number; nodeCount: number }>();

  const groupedBySheet = new Map<string, GraphNode[]>();
  for (const node of graphNodes) {
    if (!groupedBySheet.has(node.sheet)) {
      groupedBySheet.set(node.sheet, []);
    }
    groupedBySheet.get(node.sheet)?.push(node);
  }

  const orderedSheets = [...groupedBySheet.keys()].sort((a, b) => a.localeCompare(b));

  const perSheet = new Map<string, { positions: Map<string, { x: number; y: number }>; width: number; height: number; nodeCount: number }>();
  for (const sheet of orderedSheets) {
    const nodes = groupedBySheet.get(sheet) ?? [];
    const idSet = new Set(nodes.map((n) => n.id));
    const edges = graphEdges.filter((e) => idSet.has(e.source) && idSet.has(e.target));
    perSheet.set(sheet, buildDagreLayout(nodes, edges));
  }

  const rowHeights: number[] = [];
  const colWidths: number[] = [];

  if (selectedSheet === "ALL") {
    const rowCount = Math.ceil(orderedSheets.length / SHEETS_PER_ROW);
    for (let row = 0; row < rowCount; row += 1) {
      let maxHeight = 0;
      const from = row * SHEETS_PER_ROW;
      const to = Math.min(from + SHEETS_PER_ROW, orderedSheets.length);
      for (let i = from; i < to; i += 1) {
        const entry = perSheet.get(orderedSheets[i]);
        if (entry) {
          maxHeight = Math.max(maxHeight, entry.height + SHEET_PADDING_Y * 2);
        }
      }
      rowHeights[row] = maxHeight;
    }

    for (let col = 0; col < SHEETS_PER_ROW; col += 1) {
      let maxWidth = 0;
      for (let i = col; i < orderedSheets.length; i += SHEETS_PER_ROW) {
        const entry = perSheet.get(orderedSheets[i]);
        if (entry) {
          maxWidth = Math.max(maxWidth, entry.width + SHEET_PADDING_X * 2);
        }
      }
      colWidths[col] = maxWidth;
    }
  }

  orderedSheets.forEach((sheet, sheetIdx) => {
    const entry = perSheet.get(sheet);
    if (!entry) {
      return;
    }

    const col = selectedSheet === "ALL" ? sheetIdx % SHEETS_PER_ROW : 0;
    const row = selectedSheet === "ALL" ? Math.floor(sheetIdx / SHEETS_PER_ROW) : 0;

    let offsetX = 0;
    let offsetY = 0;

    if (selectedSheet === "ALL") {
      for (let c = 0; c < col; c += 1) {
        offsetX += (colWidths[c] ?? 0) + SHEET_GAP_X;
      }
      for (let r = 0; r < row; r += 1) {
        offsetY += (rowHeights[r] ?? 0) + SHEET_GAP_Y;
      }
    }

    const groupX = offsetX;
    const groupY = offsetY;
    const groupWidth = entry.width + SHEET_PADDING_X * 2;
    const groupHeight = entry.height + SHEET_PADDING_Y * 2;

    sheetBounds.set(sheet, {
      x: groupX,
      y: groupY,
      width: groupWidth,
      height: groupHeight,
      nodeCount: entry.nodeCount
    });

    for (const [id, pos] of entry.positions) {
      positions.set(id, {
        x: groupX + SHEET_PADDING_X + pos.x,
        y: groupY + SHEET_PADDING_Y + pos.y
      });
    }
  });

  return { positions, sheetBounds };
}

function buildDagreLayout(
  nodes: GraphNode[],
  edges: GraphEdge[]
): { positions: Map<string, { x: number; y: number }>; width: number; height: number; nodeCount: number } {
  const graph = new dagre.graphlib.Graph();
  graph.setDefaultEdgeLabel(() => ({}));
  graph.setGraph({
    rankdir: "LR",
    nodesep: 62,
    ranksep: 118,
    marginx: 0,
    marginy: 0
  });

  for (const node of nodes) {
    graph.setNode(node.id, { width: NODE_WIDTH, height: NODE_HEIGHT });
  }

  for (const edge of edges) {
    graph.setEdge(edge.source, edge.target);
  }

  dagre.layout(graph);

  const raw = new Map<string, { x: number; y: number }>();
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  for (const node of nodes) {
    const p = graph.node(node.id);
    const x = p.x - NODE_WIDTH / 2;
    const y = p.y - NODE_HEIGHT / 2;
    raw.set(node.id, { x, y });
    minX = Math.min(minX, x);
    minY = Math.min(minY, y);
    maxX = Math.max(maxX, x + NODE_WIDTH);
    maxY = Math.max(maxY, y + NODE_HEIGHT);
  }

  const normalized = new Map<string, { x: number; y: number }>();
  for (const [id, pos] of raw) {
    normalized.set(id, { x: pos.x - minX, y: pos.y - minY });
  }

  const width = Number.isFinite(maxX) ? Math.max(320, maxX - minX) : 320;
  const height = Number.isFinite(maxY) ? Math.max(240, maxY - minY) : 240;

  return {
    positions: normalized,
    width,
    height,
    nodeCount: nodes.length
  };
}

export function toFlowEdges(
  graphEdges: GraphEdge[],
  highlight: Set<string>,
  nodeSheetMap: Map<string, string>,
  selectedNodeId: string | null,
  hoveredNodeId: string | null
): Edge[] {
  return graphEdges.map((edge) => {
    const active = highlight.has(edge.source) || highlight.has(edge.target);
    const isCrossSheet = nodeSheetMap.get(edge.source) !== nodeSheetMap.get(edge.target);
    const hasSelection = Boolean(selectedNodeId);
    const hasHover = Boolean(hoveredNodeId);

    const stroke = active ? (isCrossSheet ? "#6d28d9" : "#4b5563") : (isCrossSheet ? "#a78bfa" : "#94a3b8");
    const opacity = hasSelection
      ? (active ? 0.92 : 0.12)
      : hasHover
        ? (active ? 0.78 : 0.14)
        : 0.28;

    return {
      id: `${edge.source}->${edge.target}`,
      source: edge.source,
      target: edge.target,
      type: "smoothstep",
      className: isCrossSheet ? "edge-cross-sheet" : "edge-same-sheet",
      markerEnd: {
        type: MarkerType.ArrowClosed,
        width: 12,
        height: 12,
        color: stroke
      },
      animated: active && (hasSelection || hasHover),
      style: {
        stroke,
        strokeDasharray: isCrossSheet ? "5 4" : undefined,
        strokeWidth: active ? 1.8 : 1.1,
        opacity,
        transition: "opacity 160ms ease, stroke 160ms ease"
      }
    } satisfies Edge;
  });
}

export function buildTraversalSets(startId: string, edges: GraphEdge[]): { upstream: Set<string>; downstream: Set<string> } {
  const upstream = new Set<string>();
  const downstream = new Set<string>();

  const incoming = new Map<string, string[]>();
  const outgoing = new Map<string, string[]>();

  for (const edge of edges) {
    if (!incoming.has(edge.target)) incoming.set(edge.target, []);
    if (!outgoing.has(edge.source)) outgoing.set(edge.source, []);
    incoming.get(edge.target)!.push(edge.source);
    outgoing.get(edge.source)!.push(edge.target);
  }

  const walk = (seed: string, map: Map<string, string[]>, result: Set<string>): void => {
    const queue = [seed];
    while (queue.length > 0) {
      const current = queue.shift();
      if (!current) continue;
      for (const next of map.get(current) ?? []) {
        if (!result.has(next)) {
          result.add(next);
          queue.push(next);
        }
      }
    }
  };

  walk(startId, incoming, upstream);
  walk(startId, outgoing, downstream);
  return { upstream, downstream };
}

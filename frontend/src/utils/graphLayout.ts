import dagre from "dagre";
import { Edge, MarkerType, Node } from "@xyflow/react";
import { CellValue, GraphEdge, GraphNode } from "../types/workbook";

const SHEET_COLORS = ["#16a34a", "#2563eb", "#9333ea", "#ea580c", "#0891b2", "#ca8a04", "#be123c"];
const ROLE_COLORS = {
  input: "#2563eb",
  computed: "#0f766e",
  output: "#16a34a",
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
const HANDLE_SLOT_COUNT = 4;

export interface FlowCellData {
  [key: string]: unknown;
  label: string;
  id: string;
  fileName: string;
  sheet: string;
  value?: CellValue;
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
  fileName: string;
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
  selectedFile: string | "ALL";
  selectedSheet: string | "ALL";
  zoomLevel: number;
}

interface LayoutResult {
  positions: Map<string, { x: number; y: number }>;
  sheetBounds: Map<
    string,
    {
      fileName: string;
      sheet: string;
      x: number;
      y: number;
      width: number;
      height: number;
      nodeCount: number;
    }
  >;
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

  const sheetGroups: Array<Node<FlowSheetGroupData>> = [...layout.sheetBounds.entries()].map(([groupKey, bounds]) => ({
    id: `group:${groupKey}`,
    type: "sheetGroup",
    position: { x: bounds.x, y: bounds.y },
    data: {
      fileName: bounds.fileName,
      sheet: bounds.sheet,
      nodeCount: bounds.nodeCount,
      color: hashSheetColor(groupKey)
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

    let role: Role = node.fileRole === "output" ? "output" : node.fileRole === "input" ? "input" : (node.formula ? "computed" : "input");
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
        fileName: node.fileName,
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
  const sheetBounds = new Map<string, { fileName: string; sheet: string; x: number; y: number; width: number; height: number; nodeCount: number }>();

  const groupedBySheet = new Map<string, GraphNode[]>();
  for (const node of graphNodes) {
    const groupKey = `${node.fileName}::${node.sheet}`;
    if (!groupedBySheet.has(groupKey)) {
      groupedBySheet.set(groupKey, []);
    }
    groupedBySheet.get(groupKey)?.push(node);
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

  orderedSheets.forEach((groupKey, sheetIdx) => {
    const entry = perSheet.get(groupKey);
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

    const [fileName, sheet] = groupKey.split("::");

    sheetBounds.set(groupKey, {
      fileName,
      sheet,
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
  const positions = new Map<string, { x: number; y: number }>();

  if (nodes.length === 0) {
    return { positions, width: 320, height: 240, nodeCount: 0 };
  }

  const cols = Math.ceil(Math.sqrt(nodes.length));
  const rows = Math.ceil(nodes.length / cols);

  const cellWidth = NODE_WIDTH + 90;
  const cellHeight = NODE_HEIGHT + 90;

  nodes.forEach((node, index) => {
    const row = Math.floor(index / cols);
    const col = index % cols;

    const x = col * cellWidth;
    const y = row * cellHeight;

    positions.set(node.id, { x, y });
  });

  const width = Math.max(320, cols * cellWidth);
  const height = Math.max(240, rows * cellHeight);

  return {
    positions,
    width,
    height,
    nodeCount: nodes.length
  };
}

export function toFlowEdges(
  graphEdges: GraphEdge[],
  highlight: Set<string>,
  nodeFileMap: Map<string, string>,
  selectedNodeId: string | null,
  hoveredNodeId: string | null
): Edge[] {
  const sortedEdges = [...graphEdges].sort((left, right) => {
    const sourceCompare = left.source.localeCompare(right.source);
    if (sourceCompare !== 0) {
      return sourceCompare;
    }
    return left.target.localeCompare(right.target);
  });

  const sourceTargets = new Map<string, string[]>();
  const targetSources = new Map<string, string[]>();

  for (const edge of sortedEdges) {
    if (!sourceTargets.has(edge.source)) {
      sourceTargets.set(edge.source, []);
    }
    sourceTargets.get(edge.source)!.push(edge.target);

    if (!targetSources.has(edge.target)) {
      targetSources.set(edge.target, []);
    }
    targetSources.get(edge.target)!.push(edge.source);
  }

  for (const [source, targets] of sourceTargets) {
    sourceTargets.set(source, [...new Set(targets)].sort((a, b) => a.localeCompare(b)));
  }

  for (const [target, sources] of targetSources) {
    targetSources.set(target, [...new Set(sources)].sort((a, b) => a.localeCompare(b)));
  }

  return sortedEdges.map((edge) => {
    const sourceIndex = sourceTargets.get(edge.source)?.indexOf(edge.target) ?? 0;
    const targetIndex = targetSources.get(edge.target)?.indexOf(edge.source) ?? 0;
    const sourceSlot = sourceIndex % HANDLE_SLOT_COUNT;
    const targetSlot = targetIndex % HANDLE_SLOT_COUNT;

    const active = selectedNodeId
      ? edge.target === selectedNodeId && highlight.has(edge.source)
      : highlight.has(edge.source) && highlight.has(edge.target);
    const isCrossFile = nodeFileMap.get(edge.source) !== nodeFileMap.get(edge.target);
    const hasSelection = Boolean(selectedNodeId);
    const hasHover = Boolean(hoveredNodeId);

    const stroke = active
      ? (isCrossFile ? "#7c3aed" : "#0f766e")
      : (isCrossFile ? "#b6a9df" : "#9ab3a5");
    const opacity = hasSelection
      ? (active ? 0.96 : 0.02)
      : hasHover
        ? (active ? 0.86 : 0.12)
        : 0.42;

    return {
      id: `${edge.source}->${edge.target}`,
      source: edge.source,
      target: edge.target,
      sourceHandle: `out-${sourceSlot % HANDLE_SLOT_COUNT}`,
      targetHandle: `in-${targetSlot % HANDLE_SLOT_COUNT}`,
      type: "bezier",
      className: isCrossFile ? "edge-cross-file" : "edge-same-sheet",
      zIndex: 0,
      markerEnd: {
        type: MarkerType.ArrowClosed,
        width: active ? 13 : 10,
        height: active ? 13 : 10,
        color: stroke
      },
      animated: false,
      style: {
        stroke,
        strokeDasharray: isCrossFile ? (active ? "7 5" : "4 4") : undefined,
        strokeWidth: active ? 2.6 : 1.55,
        strokeLinecap: "round",
        opacity,
        transition: "opacity 180ms ease, stroke 180ms ease, stroke-width 180ms ease"
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

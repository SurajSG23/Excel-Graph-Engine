import { Edge, MarkerType, Node } from "@xyflow/react";
import { GraphEdge, GraphNode } from "../types/workbook";

const SHEET_COLORS = ["#10b981", "#34d399", "#22c55e", "#16a34a", "#2dd4bf", "#4ade80", "#6ee7b7"];

export interface FlowCellData {
  [key: string]: unknown;
  label: string;
  sheet: string;
  value?: number;
  formula?: string;
  color: string;
  isSelected: boolean;
  isHighlighted: boolean;
  dependencyCount: number;
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
  highlight: Set<string>,
  selectedNodeId: string | null
): Node<FlowCellData>[] {
  const grouped = new Map<string, GraphNode[]>();
  for (const n of graphNodes) {
    if (!grouped.has(n.sheet)) {
      grouped.set(n.sheet, []);
    }
    grouped.get(n.sheet)!.push(n);
  }

  const nodes: Node<FlowCellData>[] = [];
  const sheets = [...grouped.keys()];

  sheets.forEach((sheet, sheetIndex) => {
    const entries = grouped.get(sheet) ?? [];
    entries.sort((a, b) => a.cell.localeCompare(b.cell));

    entries.forEach((node, idx) => {
      const color = hashSheetColor(node.sheet);
      const isSelected = selectedNodeId === node.id;
      const isHighlighted = highlight.has(node.id);

      nodes.push({
        id: node.id,
        type: "cellNode",
        data: {
          label: `${node.cell}`,
          sheet: node.sheet,
          value: node.value,
          formula: node.formula,
          color,
          isSelected,
          isHighlighted,
          dependencyCount: node.dependencies.length
        },
        position: {
          x: sheetIndex * 470 + (idx % 3) * 145,
          y: Math.floor(idx / 3) * 138
        }
      });
    });
  });

  return nodes;
}

export function toFlowEdges(graphEdges: GraphEdge[], highlight: Set<string>): Edge[] {
  return graphEdges.map((edge) => {
    const active = highlight.has(edge.source) || highlight.has(edge.target);
    return {
      id: `${edge.source}->${edge.target}`,
      source: edge.source,
      target: edge.target,
      type: "smoothstep",
      markerEnd: {
        type: MarkerType.ArrowClosed,
        width: 18,
        height: 18,
        color: active ? "#059669" : "#64748b"
      },
      animated: active,
      style: {
        stroke: active ? "#059669" : "#64748b",
        strokeWidth: active ? 2.4 : 1.5,
        opacity: active ? 1 : 0.55
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

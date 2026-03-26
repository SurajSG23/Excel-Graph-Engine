import { useMemo } from "react";
import { Background, Controls, Edge, MarkerType, Node, ReactFlow } from "@xyflow/react";
import "@xyflow/react/dist/style.css";
import { useWorkbookStore } from "../store/workbookStore";

function rowPlaceholder(index: number): string {
  const base = ["x", "y", "z", "u", "v", "w", "p", "q", "r", "s", "t"];
  if (index < base.length) {
    return base[index];
  }

  return `v${index - base.length + 1}`;
}

function toFormulaTemplate(formula: string): string {
  const normalized = formula.replace(/\s+/g, " ").trim();
  if (!normalized) {
    return "=f(Bx)";
  }

  const referenceMap = new Map<string, string>();
  let referenceIndex = 0;

  const templated = normalized.replace(/(\$?[A-Z]{1,3})(\$?\d+)/gi, (_, colPart: string, rowPart: string) => {
    const key = `${colPart}${rowPart}`.toUpperCase();
    if (!referenceMap.has(key)) {
      referenceMap.set(key, rowPlaceholder(referenceIndex));
      referenceIndex += 1;
    }

    const placeholder = referenceMap.get(key) ?? "x";
    const rowPrefix = rowPart.startsWith("$") ? "$" : "";
    return `${colPart}${rowPrefix}${placeholder}`;
  });

  if (templated.length <= 40) {
    return templated;
  }

  return `${templated.slice(0, 37)}...`;
}

export function GraphCanvas() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedNodeId = useWorkbookStore((s) => s.selectedNodeId);
  const setSelectedNode = useWorkbookStore((s) => s.setSelectedNode);

  const nodes = useMemo<Node[]>(() => {
    if (!workbook) {
      return [];
    }

    const formulaCount = workbook.config.formulas.length;
    const centerOffset = Math.max(0, Math.floor(formulaCount / 2));

    return workbook.graph.nodes.map((item) => {
      if (item.type === "input") {
        return {
          id: item.id,
          position: { x: 80, y: 200 },
          data: { label: item.label },
          style: {
            borderRadius: 12,
            background: "#0f4c5c",
            color: "#fff",
            border: "2px solid #073b4c",
            width: 180
          }
        };
      }

      if (item.type === "output") {
        return {
          id: item.id,
          position: { x: 700, y: 200 },
          data: { label: item.label },
          style: {
            borderRadius: 12,
            background: "#31572c",
            color: "#fff",
            border: "2px solid #132a13",
            width: 180
          }
        };
      }

      const formulaIndex = workbook.config.formulas.findIndex((node) => node.id === item.id);
      const formulaNode = workbook.config.formulas.find((node) => node.id === item.id);
      const vertical = 80 + (formulaIndex - centerOffset) * 110;
      const isSelected = selectedNodeId === item.id;
      const formulaLabel = formulaNode
        ? toFormulaTemplate(formulaNode.formula)
        : "=f(Bx)";

      return {
        id: item.id,
        position: { x: 360, y: vertical },
        data: { label: formulaLabel },
        style: {
          borderRadius: 10,
          background: isSelected ? "#f4a261" : "#e9c46a",
          color: "#1f2937",
          border: isSelected ? "2px solid #9b2226" : "1px solid #bc6c25",
          width: 280,
          fontFamily: '"IBM Plex Mono", monospace',
          fontSize: "12px",
          fontWeight: 500,
          overflow: "hidden",
          textOverflow: "ellipsis",
          whiteSpace: "nowrap",
          padding: "8px 10px"
        }
      };
    });
  }, [selectedNodeId, workbook]);

  const edges = useMemo<Edge[]>(() => {
    if (!workbook) {
      return [];
    }

    return workbook.graph.edges.map((item) => ({
      id: `${item.source}->${item.target}`,
      source: item.source,
      target: item.target,
      animated: item.source === "input",
      type: "smoothstep",
      markerEnd: {
        type: MarkerType.ArrowClosed,
        width: 18,
        height: 18,
        color: item.source === "input" ? "#0f4c5c" : "#31572c"
      },
      style: {
        stroke: item.source === "input" ? "#0f4c5c" : "#31572c",
        strokeWidth: item.source === "input" ? 2.2 : 2,
        opacity: 0.9
      }
    }));
  }, [workbook]);

  if (!workbook) {
    return (
      <section className="canvas-empty">
        <h2>No workbook loaded</h2>
        <p>Upload an Excel workbook to generate Input, Formula, and Output pipeline nodes.</p>
      </section>
    );
  }

  return (
    <section className="canvas-shell">
      <ReactFlow
        nodes={nodes}
        edges={edges}
        fitView
        minZoom={0.35}
        maxZoom={1.8}
        onNodeClick={(_, node) => setSelectedNode(node.id)}
        onPaneClick={() => setSelectedNode(null)}
      >
        <Background color="#000000" gap={28} />
        <Controls showInteractive={false} />
      </ReactFlow>
    </section>
  );
}

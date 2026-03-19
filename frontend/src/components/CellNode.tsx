import { Handle, Node, NodeProps, Position } from "@xyflow/react";
import { FlowCellData } from "../utils/graphLayout";

type CellFlowNode = Node<FlowCellData, "cellNode">;

function formatValue(value: number | undefined): string {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return "N/A";
  }

  if (Math.abs(value) >= 1000) {
    return Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(value);
  }

  return Number.isInteger(value) ? String(value) : value.toFixed(2);
}

export function CellNode({ data }: NodeProps<CellFlowNode>) {
  const nodeData = data;
  const formulaPreview = (nodeData.formula ?? "(constant value)").slice(0, 42);

  return (
    <article
      className={`cell-node ${nodeData.isHighlighted ? "is-highlighted" : ""} ${nodeData.isSelected ? "is-selected" : ""}`}
      style={{
        ["--sheet-color" as string]: nodeData.color
      }}
    >
      <Handle type="target" position={Position.Left} className="cell-node-handle" />
      <header className="cell-node-header">
        <strong>{nodeData.label}</strong>
        <span>{nodeData.sheet}</span>
      </header>

      <div className="cell-node-value">{formatValue(nodeData.value)}</div>

      <p className="cell-node-formula" title={nodeData.formula ?? "No formula"}>
        {formulaPreview}
        {formulaPreview.length < (nodeData.formula ?? "").length ? "..." : ""}
      </p>

      <footer className="cell-node-footer">
        <span>{nodeData.dependencyCount} deps</span>
      </footer>

      <Handle type="source" position={Position.Right} className="cell-node-handle" />
    </article>
  );
}

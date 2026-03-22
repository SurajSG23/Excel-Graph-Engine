import { Handle, Node, NodeProps, Position } from "@xyflow/react";
import { FlowCellData } from "../utils/graphLayout";

type CellFlowNode = Node<FlowCellData, "cellNode">;
const HANDLE_POSITIONS = [24, 40, 56, 72];

function formatValue(value: string | number | boolean | undefined): string {
  if (value === undefined) {
    return "";
  }

  if (typeof value === "string") {
    return value;
  }

  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  if (Number.isNaN(value)) {
    return "";
  }

  if (Math.abs(value) >= 1000) {
    return Intl.NumberFormat("en-US", { maximumFractionDigits: 2 }).format(value);
  }

  return Number.isInteger(value) ? String(value) : value.toFixed(2);
}

export function CellNode({ data }: NodeProps<CellFlowNode>) {
  const nodeData = data;
  const roleLabel = nodeData.role === "computed" ? "Computed" :
    nodeData.role === "output" ? "Output" :
      nodeData.role === "error" ? "Error" :
        nodeData.role === "circular" ? "Circular" : "Input";

  return (
    <article
      className={[
        "cell-node",
        `role-${nodeData.role}`,
        nodeData.isDimmed ? "is-dimmed" : "",
        nodeData.isHighlighted ? "is-highlighted" : "",
        nodeData.isHovered ? "is-hovered" : "",
        nodeData.isSelected ? "is-selected" : ""
      ].join(" ")}
      style={{
        ["--sheet-color" as string]: nodeData.color,
        ["--role-color" as string]: nodeData.roleColor
      }}
    >
      {HANDLE_POSITIONS.map((top, index) => (
        <Handle
          key={`in-${index}`}
          id={`in-${index}`}
          type="target"
          position={Position.Left}
          className="cell-node-handle"
          style={{ top: `${top}%`, transform: "translate(-50%, -50%)" }}
        />
      ))}
      <header className="cell-node-header">
        <strong>{nodeData.label}</strong>
        {nodeData.showExtra && <span>{nodeData.sheet}</span>}
      </header>

      <div className="cell-node-role">{roleLabel}</div>

      <div className="cell-node-value">{formatValue(nodeData.value)}</div>

      <footer className="cell-node-footer">
        {nodeData.showExtra && <span>{nodeData.dependencyCount} deps</span>}
        {nodeData.isUpstream && <span className="dir-chip">upstream</span>}
        {nodeData.isDownstream && <span className="dir-chip">downstream</span>}
      </footer>

      <div className="cell-node-tooltip">
        <p><strong>{nodeData.id}</strong></p>
        <p>Value: {formatValue(nodeData.value) || "available"}</p>
        <p>Formula: {nodeData.formula ?? "(none)"}</p>
        <p>Dependencies: {nodeData.dependencies.length ? nodeData.dependencies.join(", ") : "None"}</p>
      </div>

      {HANDLE_POSITIONS.map((top, index) => (
        <Handle
          key={`out-${index}`}
          id={`out-${index}`}
          type="source"
          position={Position.Right}
          className="cell-node-handle"
          style={{ top: `${top}%`, transform: "translate(50%, -50%)" }}
        />
      ))}
    </article>
  );
}

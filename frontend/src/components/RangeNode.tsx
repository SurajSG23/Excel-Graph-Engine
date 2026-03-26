import { Handle, Node, NodeProps, Position } from "@xyflow/react";
import { FlowRangeData } from "../utils/graphLayout";

type RangeFlowNode = Node<FlowRangeData, "cellNode" | "rangeNode">;

const HANDLE_POSITIONS = [25, 50, 75];

export function RangeNode({ data }: NodeProps<RangeFlowNode>) {
  return (
    <article
      className={[
        "cell-node",
        `cell-node-${data.nodeType}`,
        data.isDimmed ? "is-dimmed" : "",
        data.isHighlighted ? "is-highlighted" : "",
        data.isSelected ? "is-selected" : ""
      ].join(" ")}
      style={{
        ["--node-color" as string]: data.color,
        ["--role-color" as string]: data.roleColor
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
        <strong>{data.nodeType.toUpperCase()}</strong>
        <span>{data.operation ?? "ExcelFormula"}</span>
      </header>
      <p className="cell-node-range">{data.range}</p>
      <p className="cell-node-meta">{data.sheet}</p>

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

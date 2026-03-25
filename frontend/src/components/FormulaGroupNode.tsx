import { Handle, Node, NodeProps, Position } from "@xyflow/react";
import { FlowFormulaGroupData } from "../utils/graphLayout";

type FormulaGroupFlowNode = Node<FlowFormulaGroupData, "formulaGroup">;

const HANDLE_POSITIONS = [28, 50, 72];

export function FormulaGroupNode({ data }: NodeProps<FormulaGroupFlowNode>) {
  return (
    <article
      className={[
        "formula-group-node",
        data.isDimmed ? "is-dimmed" : "",
        data.isHighlighted ? "is-highlighted" : "",
        data.isSelected ? "is-selected" : ""
      ].join(" ")}
      style={{ ["--role-color" as string]: data.roleColor }}
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

      <header>
        <strong>Grouped Formula</strong>
        <span>{data.memberCount} cells</span>
      </header>
      <p className="formula-group-template">{data.formulaTemplate}</p>
      <footer>
        <span>{data.inputCount} inputs</span>
        <span>{data.outputCount} outputs</span>
      </footer>

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

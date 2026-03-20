import { Node, NodeProps } from "@xyflow/react";
import { FlowSheetGroupData } from "../utils/graphLayout";

type SheetGroupFlowNode = Node<FlowSheetGroupData, "sheetGroup">;

export function SheetGroupNode({ data }: NodeProps<SheetGroupFlowNode>) {
  return (
    <div className="sheet-group-node" style={{ ["--sheet-color" as string]: data.color }}>
      <div className="sheet-group-chip">
        <strong>{data.sheet}</strong>
        <span>{data.nodeCount} nodes</span>
      </div>
    </div>
  );
}

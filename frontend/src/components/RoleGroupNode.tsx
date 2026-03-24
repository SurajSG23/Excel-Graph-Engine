import { Node, NodeProps } from "@xyflow/react";
import { FlowRoleGroupData } from "../utils/graphLayout";

type RoleGroupFlowNode = Node<FlowRoleGroupData, "roleGroup">;

export function RoleGroupNode({ data }: NodeProps<RoleGroupFlowNode>) {
  return (
    <div className="role-group-node" style={{ ["--role-group-color" as string]: data.color }}>
      <div className="role-group-header">
        <span className="role-group-label">{data.label}</span>
        <span className="role-group-count">{data.sheetCount} sheets</span>
      </div>
    </div>
  );
}

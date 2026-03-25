import { FormEvent, useEffect, useMemo, useState } from "react";
import { useWorkbookStore } from "../store/workbookStore";
import { isGroupedNode, projectGraphForFormulaGrouping } from "../utils/formulaGrouping";

function formatNodeValue(value: string | number | boolean | undefined): string {
  if (value === undefined) {
    return "available";
  }

  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  return String(value);
}

export function NodeInspector() {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedNodeId = useWorkbookStore((s) => s.selectedNodeId);
  const groupSimilarFormulas = useWorkbookStore((s) => s.groupSimilarFormulas);
  const applyUpdate = useWorkbookStore((s) => s.applyUpdate);

  const projection = useMemo(
    () => projectGraphForFormulaGrouping(workbook, groupSimilarFormulas),
    [workbook, groupSimilarFormulas]
  );

  const workbookNodeById = useMemo(
    () => new Map((workbook?.nodes ?? []).map((item) => [item.id, item])),
    [workbook]
  );

  const node = useMemo(() => projection.nodes.find((item) => item.id === selectedNodeId), [selectedNodeId, projection.nodes]);

  const groupMembers = useMemo(
    () =>
      node && isGroupedNode(node)
        ? node.memberNodeIds
            .map((id) => workbookNodeById.get(id))
            .filter((item): item is NonNullable<typeof item> => Boolean(item))
        : [],
    [node, workbookNodeById]
  );

  const [formula, setFormula] = useState("");

  useEffect(() => {
    if (!node) {
      setFormula("");
      return;
    }

    if (isGroupedNode(node)) {
      setFormula(groupMembers[0]?.formula ?? "");
      return;
    }

    setFormula(node.formula ?? "");
  }, [node, groupMembers]);

  const onSubmit = async (event: FormEvent): Promise<void> => {
    event.preventDefault();
    if (!node) return;

    if (isGroupedNode(node)) {
      const updates = node.memberNodeIds.map((id) => ({ id, formula }));
      await applyUpdate(
        updates,
        `Edit grouped formula (${node.memberNodeIds.length} cells)`
      );
      return;
    }

    await applyUpdate([{ id: node.id, formula }], `Edit ${node.id}`);
  };

  if (!node) {
    return (
      <section className="panel inspector">
        <details className="panel-collapsible" open>
          <summary>Node Details</summary>
          <p>Select a node to inspect and edit formulas.</p>
        </details>
      </section>
    );
  }

  return (
    <section className="panel inspector">
      <details className="panel-collapsible" open>
        <summary>Node Details</summary>
        <dl>
          <dt>ID</dt>
          <dd>{node.id}</dd>
          <dt>File</dt>
          <dd>{node.fileName}</dd>
          <dt>Sheet</dt>
          <dd>{node.sheet}</dd>
          <dt>Cell</dt>
          <dd>{isGroupedNode(node) ? "Grouped" : node.cell}</dd>
          <dt>Computed Value</dt>
          <dd>{isGroupedNode(node) ? "Group node (multiple outputs)" : formatNodeValue(node.value)}</dd>
          {isGroupedNode(node) && (
            <>
              <dt>Formula Template</dt>
              <dd>{node.formulaTemplate}</dd>
              <dt>Group Size</dt>
              <dd>{node.memberNodeIds.length} formula cells</dd>
            </>
          )}
          <dt>Dependencies</dt>
          <dd>{node.dependencies.length > 0 ? node.dependencies.join(", ") : "None"}</dd>
          <dt>Reference Details</dt>
          <dd>
            {node.referenceDetails.length > 0
              ? node.referenceDetails
                  .map((ref) => `${ref.file}::${ref.sheet}::${ref.cell}${ref.external ? " (external)" : ""}`)
                  .join(", ")
              : "None"}
          </dd>
          {isGroupedNode(node) && (
            <>
              <dt>Outputs</dt>
              <dd>{node.outputs.join(", ")}</dd>
              <dt>Input→Output map</dt>
              <dd>
                {node.inputOutputMapping
                  .map((entry) => `${entry.outputNodeId} <= [${entry.inputNodeIds.join(", ")}]`)
                  .join("; ")}
              </dd>
            </>
          )}
        </dl>

        <form onSubmit={onSubmit} className="formula-form">
          <label htmlFor="formula">Formula</label>
          <textarea
            id="formula"
            value={formula}
            onChange={(e) => setFormula(e.target.value)}
            placeholder="=A1+B1"
            rows={4}
          />
          <button type="submit">{isGroupedNode(node) ? "Apply To Group + Recompute" : "Apply + Recompute"}</button>
        </form>

        {/* Add/move/delete/copy-paste controls removed per request */}
      </details>
    </section>
  );
}

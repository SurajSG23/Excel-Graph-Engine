import { FormEvent, useEffect, useMemo, useState } from "react";
import { useWorkbookStore } from "../store/workbookStore";

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
  const applyUpdate = useWorkbookStore((s) => s.applyUpdate);

  const node = useMemo(
    () => workbook?.nodes.find((item) => item.id === selectedNodeId),
    [selectedNodeId, workbook]
  );

  const [formula, setFormula] = useState("");

  useEffect(() => {
    setFormula(node?.formula ?? "");
  }, [node]);

  const onSubmit = async (event: FormEvent): Promise<void> => {
    event.preventDefault();
    if (!node) return;
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
          <dt>Sheet</dt>
          <dd>{node.sheet}</dd>
          <dt>Cell</dt>
          <dd>{node.cell}</dd>
          <dt>Computed Value</dt>
          <dd>{formatNodeValue(node.value)}</dd>
          <dt>Dependencies</dt>
          <dd>{node.dependencies.length > 0 ? node.dependencies.join(", ") : "None"}</dd>
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
          <button type="submit">Apply + Recompute</button>
        </form>
      </details>
    </section>
  );
}

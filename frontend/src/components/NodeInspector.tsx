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
    if (!isGroupedNode(node) && node.nodeType !== "formula") {
      return;
    }

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
          <p className="inspector-empty">Select a node to inspect and edit formulas.</p>
        </details>
      </section>
    );
  }

  const isGroup = isGroupedNode(node);
  const dependencyItems = node.dependencies;
  const referenceItems = node.referenceDetails.map((ref) => `${ref.file}::${ref.sheet}::${ref.cell}${ref.external ? " (external)" : ""}`);
  const outputItems = isGroup ? node.outputs : [];
  const mappingItems = isGroup
    ? node.inputOutputMapping.map((entry) => ({
        outputNodeId: entry.outputNodeId,
        inputNodeIds: entry.inputNodeIds,
      }))
    : [];

  return (
    <section className="panel inspector">
      <details className="panel-collapsible" open>
        <summary>Node Details</summary>
        <div className="inspector-hero">
          <div>
            <p className="inspector-kicker">Selected node</p>
            <h3>{isGroup ? node.formulaTemplate : node.cell}</h3>
            <p className="inspector-subtitle">
              {node.fileName} · {node.sheet}
            </p>
          </div>
          <span className={`inspector-badge ${isGroup ? "is-group" : "is-cell"}`}>
            {isGroup ? "Grouped" : "Cell"}
          </span>
        </div>

        <div className="inspector-meta-row">
          <span className="inspector-chip">ID: {node.id}</span>
          <span className="inspector-chip">Type: {isGroup ? "group" : node.nodeType}</span>
          <span className="inspector-chip">Range: {node.range}</span>
          <span className="inspector-chip">Operation: {node.operation ?? "none"}</span>
        </div>

        <div className="inspector-summary-grid">
          <article className="inspector-summary-card">
            <span>Computed Value</span>
            <strong>{isGroup ? "Group node" : formatNodeValue(node.value)}</strong>
          </article>
          <article className="inspector-summary-card">
            <span>Dependencies</span>
            <strong>{dependencyItems.length}</strong>
          </article>
          <article className="inspector-summary-card">
            <span>References</span>
            <strong>{referenceItems.length}</strong>
          </article>
          {isGroup && (
            <article className="inspector-summary-card inspector-summary-card-accent">
              <span>Group Size</span>
              <strong>{node.memberNodeIds.length}</strong>
            </article>
          )}
        </div>

        {isGroup && (
          <section className="inspector-section">
            <h4>Formula Template</h4>
            <code className="inspector-code-block">{node.formulaTemplate}</code>
          </section>
        )}

        <section className="inspector-section">
          <h4>Dependencies</h4>
          {dependencyItems.length > 0 ? (
            <div className="inspector-pill-list">
              {dependencyItems.map((dependency) => (
                <span key={dependency} className="inspector-pill">
                  {dependency}
                </span>
              ))}
            </div>
          ) : (
            <p className="inspector-empty">None</p>
          )}
        </section>

        <section className="inspector-section">
          <h4>Reference Details</h4>
          {referenceItems.length > 0 ? (
            <div className="inspector-list-box">
              {referenceItems.map((reference) => (
                <div key={reference} className="inspector-list-item">
                  {reference}
                </div>
              ))}
            </div>
          ) : (
            <p className="inspector-empty">None</p>
          )}
        </section>

        {isGroup && (
          <>
            <section className="inspector-section">
              <h4>Outputs</h4>
              {outputItems.length > 0 ? (
                <div className="inspector-pill-list">
                  {outputItems.map((output) => (
                    <span key={output} className="inspector-pill inspector-pill-output">
                      {output}
                    </span>
                  ))}
                </div>
              ) : (
                <p className="inspector-empty">None</p>
              )}
            </section>

            <section className="inspector-section">
              <h4>Input → Output Map</h4>
              {mappingItems.length > 0 ? (
                <div className="inspector-list-box">
                  {mappingItems.map((entry) => (
                    <div key={entry.outputNodeId} className="inspector-mapping-row">
                      <strong>{entry.outputNodeId}</strong>
                      <span>{entry.inputNodeIds.length > 0 ? entry.inputNodeIds.join(", ") : "No inputs"}</span>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="inspector-empty">None</p>
              )}
            </section>
          </>
        )}

        <form onSubmit={onSubmit} className="formula-form">
          <label htmlFor="formula">Formula</label>
          <textarea
            id="formula"
            value={formula}
            onChange={(e) => setFormula(e.target.value)}
            placeholder="=A1+B1"
            rows={4}
            disabled={!isGroupedNode(node) && node.nodeType !== "formula"}
          />
          <button type="submit" disabled={!isGroupedNode(node) && node.nodeType !== "formula"}>
            {isGroupedNode(node) ? "Apply To Group + Recompute" : "Apply + Recompute"}
          </button>
        </form>

        {/* Add/move/delete/copy-paste controls removed per request */}
      </details>
    </section>
  );
}

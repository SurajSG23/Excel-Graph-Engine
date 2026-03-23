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
  const applyOperations = useWorkbookStore((s) => s.applyOperations);

  const node = useMemo(
    () => workbook?.nodes.find((item) => item.id === selectedNodeId),
    [selectedNodeId, workbook]
  );

  const [formula, setFormula] = useState("");
  const [newCellAddress, setNewCellAddress] = useState("A1");
  const [newCellFormula, setNewCellFormula] = useState("");
  const [newCellValue, setNewCellValue] = useState("");
  const [moveTarget, setMoveTarget] = useState("");
  const [pasteTarget, setPasteTarget] = useState("A1");

  useEffect(() => {
    setFormula(node?.formula ?? "");
    setMoveTarget(node?.cell ?? "");
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
          <dt>File</dt>
          <dd>{node.fileName}</dd>
          <dt>Sheet</dt>
          <dd>{node.sheet}</dd>
          <dt>Cell</dt>
          <dd>{node.cell}</dd>
          <dt>Computed Value</dt>
          <dd>{formatNodeValue(node.value)}</dd>
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

        <form
          className="formula-form"
          onSubmit={async (event) => {
            event.preventDefault();
            await applyOperations(
              [
                {
                  type: "ADD_CELL",
                  fileName: node.fileName,
                  sheet: node.sheet,
                  cell: newCellAddress,
                  formula: newCellFormula || undefined,
                  value: newCellValue ? Number(newCellValue) || newCellValue : undefined,
                  fileRole: node.fileRole
                }
              ],
              `Add ${node.fileName}::${node.sheet}::${newCellAddress}`
            );
          }}
        >
          <label htmlFor="new-cell-address">Add Cell Address</label>
          <input
            id="new-cell-address"
            value={newCellAddress}
            onChange={(e) => setNewCellAddress(e.target.value)}
            placeholder="A11"
          />
          <label htmlFor="new-cell-formula">Formula (optional)</label>
          <input
            id="new-cell-formula"
            value={newCellFormula}
            onChange={(e) => setNewCellFormula(e.target.value)}
            placeholder="=A1+B1"
          />
          <label htmlFor="new-cell-value">Value (optional)</label>
          <input
            id="new-cell-value"
            value={newCellValue}
            onChange={(e) => setNewCellValue(e.target.value)}
            placeholder="42"
          />
          <button type="submit">Add Cell</button>
        </form>

        <div className="toolbar-group">
          <label>
            Move Selected Cell To
            <input value={moveTarget} onChange={(e) => setMoveTarget(e.target.value)} placeholder="B5" />
          </label>
          <button
            type="button"
            onClick={() =>
              applyOperations(
                [
                  {
                    type: "MOVE_CELL",
                    fromNodeId: node.id,
                    toFileName: node.fileName,
                    toSheet: node.sheet,
                    toCell: moveTarget
                  }
                ],
                `Move ${node.id}`
              )
            }
          >
            Move Cell
          </button>
          <button
            type="button"
            onClick={() => applyOperations([{ type: "DELETE_CELLS", nodeIds: [node.id] }], `Delete ${node.id}`)}
          >
            Delete Cell
          </button>
        </div>

        <div className="toolbar-group">
          <label>
            Paste Anchor
            <input value={pasteTarget} onChange={(e) => setPasteTarget(e.target.value)} placeholder="D10" />
          </label>
          <button
            type="button"
            onClick={() =>
              applyOperations(
                [
                  {
                    type: "COPY_PASTE",
                    sourceNodeIds: [node.id],
                    targetFileName: node.fileName,
                    targetSheet: node.sheet,
                    targetAnchorCell: pasteTarget
                  }
                ],
                `Copy ${node.id}`
              )
            }
          >
            Copy/Paste Node
          </button>
        </div>
      </details>
    </section>
  );
}

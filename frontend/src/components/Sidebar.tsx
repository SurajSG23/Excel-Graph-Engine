import { FormEvent, useEffect, useMemo, useState } from "react";
import { UploadCard } from "./UploadCard";
import { useWorkbookStore } from "../store/workbookStore";
import { PipelineRange } from "../types/workbook";

interface SidebarProps {
  isOpen: boolean;
  onToggle: () => void;
  onResizeStart: () => void;
}

function rangesToText(ranges: PipelineRange[]): string {
  return ranges.map((item) => `${item.sheet}!${item.range}`).join("\n");
}

function textToRanges(value: string): PipelineRange[] {
  return value
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [sheet, range] = line.includes("!")
        ? (line.split("!") as [string, string])
        : (["Sheet1", line] as [string, string]);
      return { sheet: sheet.trim(), range: range.trim().toUpperCase() };
    });
}

function displayWorkbookName(pathValue: string): string {
  const normalized = pathValue.replace(/\\/g, "/");
  const fileName = normalized.split("/").pop() ?? pathValue;
  // Multer stores uploads as `<timestamp>-OriginalName.xlsx`; show only the original name.
  return fileName.replace(/^\d+-/, "");
}

function columnToNumber(col: string): number {
  let total = 0;
  for (const ch of col.toUpperCase()) {
    total = total * 26 + (ch.charCodeAt(0) - 64);
  }
  return total;
}

function numberToColumn(value: number): string {
  let n = value;
  let col = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(65 + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

function parseCellAddress(cell: string): { col: number; row: number } | null {
  const match = cell.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
  if (!match) {
    return null;
  }
  return {
    col: columnToNumber(match[1]),
    row: Number(match[2])
  };
}

function translateFormulaForOutput(formula: string, anchorCell: string, outputCell: string, outputSheet: string): string {
  const anchor = parseCellAddress(anchorCell);
  const current = parseCellAddress(outputCell);
  if (!anchor || !current) {
    return formula;
  }

  const dRow = current.row - anchor.row;
  const dCol = current.col - anchor.col;
  return formula.replace(
    /((?:'[^']+'|[A-Za-z0-9_\.]+)!|)(\$?)([A-Z]{1,3})(\$?)([0-9]+)/g,
    (raw, sheetPrefix: string, absCol: string, colText: string, absRow: string, rowText: string) => {
      const base = parseCellAddress(`${colText}${rowText}`);
      if (!base) {
        return raw;
      }

      const nextCol = absCol ? base.col : base.col + dCol;
      const nextRow = absRow ? base.row : base.row + dRow;
      if (nextCol < 1 || nextRow < 1) {
        return raw;
      }

      const nextCell = `${numberToColumn(nextCol)}${nextRow}`;
      return `${sheetPrefix || `${outputSheet}!`}${nextCell}`;
    }
  );
}

function estimateRangeCellCount(rangeText: string): number {
  const ranges = rangeText
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  let total = 0;
  for (const segment of ranges) {
    const [left, right] = segment.includes(":")
      ? (segment.split(":") as [string, string])
      : ([segment, segment] as [string, string]);
    const leftMatch = left.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
    const rightMatch = right.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
    if (!leftMatch || !rightMatch) {
      continue;
    }

    const leftCol = columnToNumber(leftMatch[1]);
    const rightCol = columnToNumber(rightMatch[1]);
    const leftRow = Number(leftMatch[2]);
    const rightRow = Number(rightMatch[2]);
    const colCount = Math.abs(rightCol - leftCol) + 1;
    const rowCount = Math.abs(rightRow - leftRow) + 1;
    total += colCount * rowCount;
  }

  return total;
}

function expandRangeCells(rangeText: string): string[] {
  const ranges = rangeText
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);

  const out: string[] = [];
  for (const segment of ranges) {
    const [left, right] = segment.includes(":")
      ? (segment.split(":") as [string, string])
      : ([segment, segment] as [string, string]);
    const leftMatch = left.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
    const rightMatch = right.toUpperCase().match(/^([A-Z]{1,3})([0-9]+)$/);
    if (!leftMatch || !rightMatch) {
      continue;
    }

    const leftCol = columnToNumber(leftMatch[1]);
    const rightCol = columnToNumber(rightMatch[1]);
    const leftRow = Number(leftMatch[2]);
    const rightRow = Number(rightMatch[2]);

    const minCol = Math.min(leftCol, rightCol);
    const maxCol = Math.max(leftCol, rightCol);
    const minRow = Math.min(leftRow, rightRow);
    const maxRow = Math.max(leftRow, rightRow);

    for (let row = minRow; row <= maxRow; row += 1) {
      for (let col = minCol; col <= maxCol; col += 1) {
        out.push(`${numberToColumn(col)}${row}`);
      }
    }
  }

  return out;
}

export function Sidebar({ isOpen, onToggle, onResizeStart }: SidebarProps) {
  const workbook = useWorkbookStore((s) => s.workbook);
  const selectedNodeId = useWorkbookStore((s) => s.selectedNodeId);
  const loading = useWorkbookStore((s) => s.loading);
  const error = useWorkbookStore((s) => s.error);
  const undo = useWorkbookStore((s) => s.undo);
  const redo = useWorkbookStore((s) => s.redo);
  const triggerExport = useWorkbookStore((s) => s.triggerExport);
  const updateFormulaNode = useWorkbookStore((s) => s.updateFormulaNode);

  const formulaNode = useMemo(
    () => workbook?.config.formulas.find((node) => node.id === selectedNodeId) ?? null,
    [selectedNodeId, workbook]
  );

  const inputStats = useMemo(() => {
    if (!workbook) {
      return {
        totalRanges: 0,
        estimatedCellsRead: 0,
        downstreamFormulaCount: 0
      };
    }

    const totalRanges = workbook.config.input.ranges.length;
    const estimatedCellsRead = workbook.config.input.ranges.reduce(
      (sum, item) => sum + estimateRangeCellCount(item.range),
      0
    );

    return {
      totalRanges,
      estimatedCellsRead,
      downstreamFormulaCount: workbook.config.formulas.length
    };
  }, [workbook]);

  const inputSheetDetails = useMemo(() => {
    if (!workbook) {
      return [] as Array<{
        sheet: string;
        ranges: Array<{ range: string; count: number; preview: string[]; moreCount: number }>;
        totalCells: number;
      }>;
    }

    const bySheet = new Map<
      string,
      Array<{ range: string; count: number; preview: string[]; moreCount: number }>
    >();

    for (const item of workbook.config.input.ranges) {
      const cells = expandRangeCells(item.range);
      const preview = cells;
      if (!bySheet.has(item.sheet)) {
        bySheet.set(item.sheet, []);
      }
      bySheet.get(item.sheet)!.push({
        range: item.range,
        count: cells.length,
        preview,
        moreCount: 0
      });
    }

    return [...bySheet.entries()].map(([sheet, ranges]) => ({
      sheet,
      ranges,
      totalCells: ranges.reduce((sum, rangeInfo) => sum + rangeInfo.count, 0)
    }));
  }, [workbook]);

  const outputSheetDetails = useMemo(() => {
    if (!workbook) {
      return [] as Array<{
        sheet: string;
        ranges: Array<{ range: string; count: number; preview: string[]; moreCount: number }>;
        totalCells: number;
      }>;
    }

    const bySheet = new Map<
      string,
      Array<{ range: string; count: number; preview: string[]; moreCount: number }>
    >();

    for (const item of workbook.config.output.ranges) {
      const cells = expandRangeCells(item.range);
      const preview = cells;
      if (!bySheet.has(item.sheet)) {
        bySheet.set(item.sheet, []);
      }
      bySheet.get(item.sheet)!.push({
        range: item.range,
        count: cells.length,
        preview,
        moreCount: 0
      });
    }

    return [...bySheet.entries()].map(([sheet, ranges]) => ({
      sheet,
      ranges,
      totalCells: ranges.reduce((sum, rangeInfo) => sum + rangeInfo.count, 0)
    }));
  }, [workbook]);

  const outputStats = useMemo(() => {
    if (!workbook) {
      return {
        totalRanges: 0,
        estimatedCellsWritten: 0,
        sourceFormulaCount: 0
      };
    }

    return {
      totalRanges: workbook.config.output.ranges.length,
      estimatedCellsWritten: workbook.config.output.ranges.reduce(
        (sum, item) => sum + estimateRangeCellCount(item.range),
        0
      ),
      sourceFormulaCount: workbook.config.formulas.length
    };
  }, [workbook]);

  const outputDependencyRows = useMemo(() => {
    if (!workbook) {
      return [] as Array<{
        outputCell: string;
        formulaName: string;
        dependsOn: string;
        expression: string;
        result: string;
      }>;
    }

    const formulaNameById = new Map(workbook.config.formulas.map((item) => [item.id, item.name]));

    return workbook.config.formulas.flatMap((node) => {
      const values = workbook.nodeResults[node.id] ?? [];
      const outputCells = node.outputCells.length > 0 ? node.outputCells : expandRangeCells(node.output.range);
      const upstreamFormulaNames = workbook.graph.edges
        .filter((edge) => edge.target === node.id && edge.source !== "input" && edge.source !== "output")
        .map((edge) => formulaNameById.get(edge.source) ?? edge.source);

      const inputRangesText = node.inputs.map((item) => `${item.sheet}!${item.range}`).join(", ");
      const upstreamText = upstreamFormulaNames.length > 0 ? upstreamFormulaNames.join(", ") : "None";
      const dependsOn = `Inputs: ${inputRangesText || "None"}; Upstream formulas: ${upstreamText}`;

      return outputCells.map((cell, index) => {
        const translated = translateFormulaForOutput(node.formula, node.anchorCell, cell, node.output.sheet);
        return {
          outputCell: `${node.output.sheet}!${cell}`,
          formulaName: node.name,
          dependsOn,
          expression: translated.startsWith("=") ? translated.slice(1) : translated,
          result: values[index] === undefined ? "-" : String(values[index])
        };
      });
    });
  }, [workbook]);

  const formulaOutputCells = useMemo(() => {
    if (!formulaNode) {
      return [] as string[];
    }

    if (formulaNode.outputCells.length > 0) {
      return formulaNode.outputCells;
    }

    return expandRangeCells(formulaNode.output.range);
  }, [formulaNode]);

  const formulaInputDetails = useMemo(() => {
    if (!formulaNode) {
      return [] as Array<{ sheet: string; range: string; count: number; startCell: string; endCell: string }>;
    }

    return formulaNode.inputs.map((item) => {
      const cells = expandRangeCells(item.range);
      return {
        sheet: item.sheet,
        range: item.range,
        count: cells.length,
        startCell: cells[0] ?? "-",
        endCell: cells[cells.length - 1] ?? "-"
      };
    });
  }, [formulaNode]);

  const formulaComputedRows = useMemo(() => {
    if (!formulaNode || !workbook) {
      return [] as Array<{ outputCell: string; expression: string; result: string }>;
    }

    const values = workbook.nodeResults[formulaNode.id] ?? [];

    return formulaOutputCells.map((cell, index) => {
      const translated = translateFormulaForOutput(
        formulaNode.formula,
        formulaNode.anchorCell,
        cell,
        formulaNode.output.sheet
      );
      return {
        outputCell: `${formulaNode.output.sheet}!${cell}`,
        expression: translated.startsWith("=") ? translated.slice(1) : translated,
        result: values[index] === undefined ? "-" : String(values[index])
      };
    });
  }, [formulaNode, formulaOutputCells, workbook]);

  const formulaExecutionIndex = useMemo(() => {
    if (!workbook || !formulaNode) {
      return -1;
    }
    return workbook.executionOrder.indexOf(formulaNode.id);
  }, [formulaNode, workbook]);

  const inputEquationGroups = useMemo(() => {
    if (!workbook) {
      return [] as Array<{
        formulaId: string;
        formulaName: string;
        inputRanges: string[];
        equations: string[];
      }>;
    }

    return workbook.config.formulas.map((node) => {
      const outputCells = node.outputCells.length > 0 ? node.outputCells : expandRangeCells(node.output.range);
      const equations = outputCells.map((cell) => {
        const translated = translateFormulaForOutput(node.formula, node.anchorCell, cell, node.output.sheet);
        return `${node.output.sheet}!${cell} = ${translated.startsWith("=") ? translated.slice(1) : translated}`;
      });

      return {
        formulaId: node.id,
        formulaName: node.name,
        inputRanges: node.inputs.map((item) => `${item.sheet}!${item.range}`),
        equations
      };
    });
  }, [workbook]);

  const [formulaText, setFormulaText] = useState("");
  const [inputRangesText, setInputRangesText] = useState("");
  const [outputSheetText, setOutputSheetText] = useState("");
  const [outputRangeText, setOutputRangeText] = useState("");
  const [formError, setFormError] = useState<string | null>(null);

  useEffect(() => {
    if (!formulaNode) {
      setFormulaText("");
      setInputRangesText("");
      setOutputRangeText("");
      setFormError(null);
      return;
    }

    setFormulaText(formulaNode.formula);
    setInputRangesText(rangesToText(formulaNode.inputs));
    setOutputSheetText(formulaNode.output.sheet);
    setOutputRangeText(formulaNode.output.range);
    setFormError(null);
  }, [formulaNode]);

  const onSubmit = async (event: FormEvent): Promise<void> => {
    event.preventDefault();
    if (!formulaNode) {
      return;
    }

    const formula = formulaText.trim();
    if (!formula.startsWith("=")) {
      setFormError("Formula must start with '='.");
      return;
    }

    const inputs = textToRanges(inputRangesText);
    const sheet = outputSheetText.trim();
    const range = outputRangeText.trim();

    if (!range || !sheet) {
      setFormError("Output range must be in Sheet!A1:B3 format.");
      return;
    }

    setFormError(null);
    await updateFormulaNode({
      id: formulaNode.id,
      formula,
      inputs,
      output: {
        sheet: sheet.trim(),
        range: range.trim().toUpperCase()
      }
    });
  };

  return (
    <aside className={`dashboard-sidebar ${isOpen ? "is-open" : "is-closed"}`}>
      <div className="sidebar-header">
        <h1>Excel Pipeline Engine</h1>
        <button
          type="button"
          className="sidebar-toggle icon-button"
          onClick={onToggle}
          aria-label={isOpen ? "Collapse controls panel" : "Expand controls panel"}
        >
          <span className={isOpen ? "icon-chevron-left" : "icon-chevron-right"} aria-hidden="true" />
        </button>
      </div>

      <div className="sidebar-content">
        <UploadCard />

        <section className="panel sidebar-actions-panel">
          <h3>Pipeline Actions</h3>
          <div className="sidebar-actions-grid">
            <button type="button" className="sidebar-action-btn" onClick={() => void undo()} disabled={loading || !workbook}>Undo</button>
            <button type="button" className="sidebar-action-btn" onClick={() => void redo()} disabled={loading || !workbook}>Redo</button>
            <button type="button" className="sidebar-action-btn sidebar-action-btn-primary" onClick={() => void triggerExport()} disabled={loading || !workbook}>Export</button>
          </div>
        </section>

        {error && <div className="error-banner">{error}</div>}
        {loading && <div className="loading-banner">Executing pipeline...</div>}

        <section className="panel inspector">
          <h3>Node Sidebar</h3>
          <p className="inspector-subtitle">Select Input, Formula, or Output node to inspect and edit pipeline behavior.</p>
          {!workbook && <p className="node-empty-state">Upload a workbook to inspect pipeline nodes.</p>}
          {workbook && selectedNodeId === "input" && (
            <div className="node-details node-details-input">
              <h4>Input Node Details</h4>
              <p className="node-details-lead">
                This node reads workbook data from configured sheets and ranges, then makes those values available to
                all formula nodes in the pipeline.
              </p>

              <div className="node-detail-metrics">
                <div>
                  <span>Source Sheets</span>
                  <strong>{workbook.config.input.sheets.length}</strong>
                </div>
                <div>
                  <span>Input Ranges</span>
                  <strong>{inputStats.totalRanges}</strong>
                </div>
                <div>
                  <span>Estimated Cells Read</span>
                  <strong>{inputStats.estimatedCellsRead}</strong>
                </div>
                <div>
                  <span>Downstream Formula Nodes</span>
                  <strong>{inputStats.downstreamFormulaCount}</strong>
                </div>
              </div>

              <div className="node-detail-block">
                <h5>Source Workbook</h5>
                <p>{displayWorkbookName(workbook.config.input.filePath)}</p>
              </div>

              <div className="node-detail-block">
                <h5>Sheets In Scope</h5>
                <ul>
                  {workbook.config.input.sheets.map((sheet) => (
                    <li key={sheet}>{sheet}</li>
                  ))}
                </ul>
              </div>

              <div className="node-detail-block">
                <h5>Input Range Summary</h5>
                <div className="node-sheet-ranges">
                  {inputSheetDetails.map((sheetInfo) => (
                    <section key={sheetInfo.sheet} className="node-range-card">
                      <div className="node-range-title-row">
                        <span>{sheetInfo.sheet}</span>
                        <span>{sheetInfo.totalCells} cells</span>
                      </div>
                      <ul className="node-range-summary-list">
                        {sheetInfo.ranges.map((rangeInfo, idx) => (
                          <li key={`${sheetInfo.sheet}:${rangeInfo.range}:${idx}`}>
                            <code>{rangeInfo.range}</code>
                            <span>{rangeInfo.count} cells</span>
                          </li>
                        ))}
                      </ul>
                    </section>
                  ))}
                </div>
              </div>

              <div className="node-detail-block">
                <h5>Input to Formula Relationships</h5>
                <p>Readable mapping of how input references are used to compute output cells.</p>
                <div className="node-equation-groups">
                  {inputEquationGroups.map((group) => (
                    <article key={group.formulaId} className="node-equation-group">
                      <header>
                        <strong>{group.formulaName}</strong>
                        <span>{group.equations.length} outputs</span>
                      </header>
                      <p className="node-equation-inputs">Inputs: {group.inputRanges.join(", ")}</p>
                      <div className="node-equation-list">
                        {group.equations.map((equation) => (
                          <code key={`${group.formulaId}:${equation}`} className="node-equation-row">{equation}</code>
                        ))}
                      </div>
                    </article>
                  ))}
                </div>
              </div>

              <div className="node-detail-block">
                <h5>Execution Order Preview</h5>
                <p>
                  {workbook.executionOrder.length > 0
                    ? workbook.executionOrder.join(" -> ")
                    : "No formula nodes available."}
                </p>
              </div>
            </div>
          )}

          {workbook && selectedNodeId === "output" && (
            <div className="node-details node-details-output">
              <h4>Output Node Details</h4>
              <p className="node-details-lead">
                This node collects computed values from formula nodes and writes them to configured ranges in the
                target workbook.
              </p>

              <div className="node-detail-metrics">
                <div>
                  <span>Target Sheets</span>
                  <strong>{outputSheetDetails.length}</strong>
                </div>
                <div>
                  <span>Output Ranges</span>
                  <strong>{outputStats.totalRanges}</strong>
                </div>
                <div>
                  <span>Estimated Cells Written</span>
                  <strong>{outputStats.estimatedCellsWritten}</strong>
                </div>
                <div>
                  <span>Source Formula Nodes</span>
                  <strong>{outputStats.sourceFormulaCount}</strong>
                </div>
              </div>

              <div className="node-detail-block">
                <h5>Target Workbook</h5>
                <p>{displayWorkbookName(workbook.config.output.targetFilePath)}</p>
              </div>

              <div className="node-detail-block">
                <h5>Output Cell Coverage</h5>
                <p>Summary of where final pipeline outputs are written.</p>
                <div className="node-result-table-wrap">
                  <table className="node-result-table">
                    <thead>
                      <tr>
                        <th>Sheet</th>
                        <th>Range</th>
                        <th>Cells</th>
                        <th>Span</th>
                      </tr>
                    </thead>
                    <tbody>
                      {outputSheetDetails.flatMap((sheetInfo) =>
                        sheetInfo.ranges.map((rangeInfo, idx) => {
                          const first = rangeInfo.preview[0] ?? "-";
                          const last = rangeInfo.preview[rangeInfo.preview.length - 1] ?? "-";
                          return (
                            <tr key={`${sheetInfo.sheet}:${rangeInfo.range}:${idx}`}>
                              <td>{sheetInfo.sheet}</td>
                              <td><code>{rangeInfo.range}</code></td>
                              <td>{rangeInfo.count}</td>
                              <td>{first} to {last}</td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="node-detail-block">
                <h5>Output Dependency Mapping</h5>
                <p>Each output cell with its producer, dependencies, and latest computed value.</p>
                <div className="node-result-table-wrap">
                  <table className="node-result-table">
                    <thead>
                      <tr>
                        <th>Output Cell</th>
                        <th>Produced By</th>
                        <th>Depends On</th>
                        <th>Expression</th>
                        <th>Result</th>
                      </tr>
                    </thead>
                    <tbody>
                      {outputDependencyRows.map((row) => (
                        <tr key={`output-map:${row.outputCell}:${row.formulaName}`}>
                          <td>{row.outputCell}</td>
                          <td>{row.formulaName}</td>
                          <td>{row.dependsOn}</td>
                          <td><code>{row.expression}</code></td>
                          <td className="node-result-value">{row.result}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {formulaNode && (
            <div className="node-details node-details-formula">
              <h4>Formula Node Details</h4>
              {/* <p className="node-details-lead">
                This node transforms input range values into output range values using the configured formula pattern.
              </p>

              <div className="node-detail-metrics">
                <div>
                  <span>Node Name</span>
                  <strong>{formulaNode.name}</strong>
                </div>
                <div>
                  <span>Execution Step</span>
                  <strong>{formulaExecutionIndex >= 0 ? `${formulaExecutionIndex + 1}/${workbook?.executionOrder.length}` : "N/A"}</strong>
                </div>
                <div>
                  <span>Input Ranges</span>
                  <strong>{formulaNode.inputs.length}</strong>
                </div>
                <div>
                  <span>Output Cells</span>
                  <strong>{formulaNode.outputCells.length}</strong>
                </div>
              </div> */}

              <div className="node-detail-block">
                <h5>Current Formula</h5>
                <p className="node-formula-logic">{formulaNode.formula}</p>
              </div>

              <div className="node-detail-block">
                <h5>Input Cell Coverage</h5>
                <div className="node-result-table-wrap">
                  <table className="node-result-table">
                    <thead>
                      <tr>
                        <th>Sheet</th>
                        <th>Range</th>
                        <th>Cells</th>
                        <th>Span</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formulaInputDetails.map((item, idx) => (
                        <tr key={`${item.sheet}:${item.range}:${idx}`}>
                          <td>{item.sheet}</td>
                          <td><code>{item.range}</code></td>
                          <td>{item.count}</td>
                          <td>{item.startCell} to {item.endCell}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="node-detail-block">
                <h5>Output Destination</h5>
                <p className="node-output-summary">
                  {formulaNode.output.sheet}!{formulaNode.output.range} ({formulaOutputCells.length} cells)
                </p>
                <p className="node-output-span">
                  Span: {formulaOutputCells[0] ?? "-"} to {formulaOutputCells[formulaOutputCells.length - 1] ?? "-"}
                </p>
              </div>

              <div className="node-detail-block">
                <h5>Latest Computed Results</h5>
                <p>Per-cell computation mapping after latest pipeline execution.</p>
                <div className="node-result-table-wrap">
                  <table className="node-result-table">
                    <thead>
                      <tr>
                        <th>Output Cell</th>
                        <th>Computed Expression</th>
                        <th>Result</th>
                      </tr>
                    </thead>
                    <tbody>
                      {formulaComputedRows.map((row) => (
                        <tr key={`${formulaNode.id}:${row.outputCell}`}>
                          <td>{row.outputCell}</td>
                          <td><code>{row.expression}</code></td>
                          <td className="node-result-value">{row.result}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>

              <form className="formula-form" onSubmit={(event) => void onSubmit(event)}>
                <h5>Edit Formula Node</h5>
                <label htmlFor="formula-expression">Formula expression</label>
                <textarea
                  id="formula-expression"
                  rows={4}
                  value={formulaText}
                  onChange={(event) => setFormulaText(event.target.value)}
                />

                <label htmlFor="formula-inputs">Input ranges (one per line, Sheet!Range)</label>
                <textarea
                  id="formula-inputs"
                  rows={5}
                  value={inputRangesText}
                  onChange={(event) => setInputRangesText(event.target.value)}
                />

                <label>Output target</label>
                <div className="range-editor-grid">
                  <div className="range-editor-field">
                    <label htmlFor="formula-output-sheet">Sheet</label>
                    <input
                      id="formula-output-sheet"
                      value={outputSheetText}
                      onChange={(event) => setOutputSheetText(event.target.value)}
                      placeholder="Summary"
                    />
                  </div>
                  <div className="range-editor-field">
                    <label htmlFor="formula-output-range">Range</label>
                    <input
                      id="formula-output-range"
                      value={outputRangeText}
                      onChange={(event) => setOutputRangeText(event.target.value.toUpperCase())}
                      placeholder="A1:B20"
                    />
                  </div>
                </div>

                {formError && <div className="error-banner">{formError}</div>}
                <button type="submit" disabled={loading}>Validate + Recompute</button>
              </form>
            </div>
          )}

          {workbook && !selectedNodeId && <p className="node-empty-state">Select Input, a Formula node, or Output in the graph.</p>}
        </section>

        {workbook && (
          <section className="panel validation-panel">
            <h3>Validation</h3>
            {workbook.validationIssues.length === 0 ? (
              <p className="validation-empty">No validation issues.</p>
            ) : (
              <ul className="validation-list">
                {workbook.validationIssues.map((issue, index) => (
                  <li key={`${issue.type}:${issue.nodeId ?? "global"}:${index}`} className="validation-item">
                    <span className="validation-badge">{issue.type}</span>
                    <span>{issue.message}</span>
                  </li>
                ))}
              </ul>
            )}
          </section>
        )}
      </div>
      <button
        type="button"
        className="sidebar-resize-handle"
        onPointerDown={(event) => {
          if (event.pointerType === "mouse" || event.pointerType === "pen") {
            event.preventDefault();
          }
          onResizeStart();
        }}
        aria-label="Resize sidebar"
      />
    </aside>
  );
}
